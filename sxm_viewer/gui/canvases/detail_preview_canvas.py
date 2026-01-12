"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import io
import itertools
import json
import math

import numpy as np
from matplotlib import patches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter
from matplotlib.widgets import RectangleSelector
from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
from mpl_toolkits.axes_grid1 import make_axes_locatable
import matplotlib
from matplotlib.collections import LineCollection
import matplotlib.patheffects as PathEffects

from ..._shared import QtCore, QtGui, QtWidgets
from .molecular_overlay import Molecule, MoleculePropertiesDialog, get_cpk_color
from ..thumbnail_render import sample_array_value, array_to_qimage

class MultiPreviewCanvas(FigureCanvas):
    def __init__(self, parent=None, figsize=(6,6)):
        self.fig = Figure(figsize=figsize)
        super().__init__(self.fig)
        if parent is not None:
            self.setParent(parent)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.views = []
        self._ax_view_map = {}
        self._copy_feedback_handler = None
        self._views_callback = None
        self._drag_candidate = None  # (view, QPoint start, QImage cache)
        self._value_callback = None
        self._value_cid = self.mpl_connect('motion_notify_event', self._on_motion_value)
        self.mpl_connect('motion_notify_event', self._on_molecule_motion)
        self.mpl_connect('button_release_event', self._on_molecule_release)
        # profile (interactive line) state
        self.profile_enabled = False
        self.profile_pts = None  # (x0, y0, x1, y1) in data coords of main ax
        self._profile_line = None
        self._profile_p0 = None
        self._profile_p1 = None
        self._profile_ticks = None
        self._profile_info_text = None
        self._profile_label = None
        self._profile_endpoint_labels = []
        self._profile_hud_text = None
        self._profile_marker_positions = None
        self._profile_marker_domain = None
        self._profile_marker_artists = []
        self._profile_marker_drag_idx = None
        self._profile_marker_callback = None
        self._profile_marker_key = None
        self._profile_marker_positions_by_key = {}
        self._profile_marker_domain_by_key = {}
        self._profile_state_callback = None
        self._profile_state_syncing = False
        self._profile_state_deferred = False
        self._profile_update_timer = QtCore.QTimer(self)
        self._profile_update_timer.setSingleShot(True)
        self._profile_update_timer.setInterval(50)
        self._profile_update_timer.timeout.connect(self._flush_profile_updates)
        self._saved_profiles = []
        self._profile_color_cycle = itertools.cycle([
            '#000000', '#e6194B', '#4363d8', '#3cb44b',
            '#911eb4', '#f58231', '#a9a9a9'
        ])
        self._line_drag_origin = None
        self._active_profile_color = '#fbc02d'
        self._active_profile_lw = 2.0
        self._highlighted_overlay = None
        self._cids = []
        self._base_click_cid = self.mpl_connect('button_press_event', self._on_base_click)
        self._dragging = None  # 'p0' or 'p1'
        self.main_ax = None
        self.profile_callback = None  # callable(active_dataset, saved_datasets)
        self._profile_highlight_cb = None
        self._profile_label_scale = 1.0
        self._view_font_scale = 1.0
        self._colorbar_orientation = 'vertical'
        self._show_ticks = True
        self._show_colorbar = True
        self._detail_dark = False
        self._detail_grid = False
        self._colorbars = []
        self._view_layout = "grid"
        self.angle_enabled = False
        self.angle_pts = None  # (vx, vy, ax, ay, bx, by)
        self._angle_lines = []
        self._angle_markers = []
        self._angle_label = None
        self._angle_len_labels = []
        self._angle_patch = None
        self._angle_dragging = None
        self._angle_cids = []
        self._angle_drag_origin = None
        self.angle_callback = None
        self.scale_bar_enabled = False
        self._scale_bar_pos = (0.94, 0.06)  # default lower right (axes coords)
        self._scale_bar_artists = []
        self._scale_bar_cids = []
        self._scale_bar_drag_start = None
        self._profile_echo_artists = []
        self._scale_bar_settings = {
            'text_color': None,
            'bar_color': None,
            'font_family': 'sans-serif'
        }
        # Molecular overlay state
        self.molecules = []
        self._molecule_drag_idx = None
        self._molecule_drag_start = None
        self._molecule_drag_start_px = None
        self._molecule_drag_mol_start = None
        self._molecule_drag_mol_angles = None
        self._molecule_drag_mode = None
        self._molecule_rotation_guide = None
        self._molecule_artists = []
        self._profile_background = None
        self._active_profile_original_color = None

    def draw(self):
        try:
            super().draw()
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return

    def set_views(self, views, preserve_profiles: bool = False):
        state = None
        if preserve_profiles:
            try:
                state = self.export_profile_state()
            except Exception:
                state = None
        self.views = views[:]
        if not preserve_profiles:
            # whenever a new view set arrives, clear saved overlays so we don't mix files
            self._clear_saved_profile_artists(notify=False)
            self.profile_pts = None
        self._redraw()
        if preserve_profiles and state is not None:
            try:
                self.import_profile_state(state, emit=False)
            except Exception:
                pass
        if callable(self._views_callback):
            try:
                self._views_callback(self.views)
            except Exception:
                pass

    def set_view_layout(self, layout: str):
        layout = (layout or "").strip().lower()
        if layout not in ("grid", "stacked"):
            layout = "grid"
        if layout == self._view_layout:
            return
        self._view_layout = layout
        self._redraw()

    def clear_views(self):
        self.views = []
        self._redraw()

    def set_views_callback(self, cb):
        self._views_callback = cb

    def resizeEvent(self, event):
        size = event.size()
        if size.width() <= 0 or size.height() <= 0:
            safe = QtCore.QSize(max(1, size.width()), max(1, size.height()))
            safe_event = QtGui.QResizeEvent(safe, event.oldSize())
            super().resizeEvent(safe_event)
            return
        super().resizeEvent(event)

    def _redraw(self):
        self.fig.clf()
        self._ax_view_map = {}
        self._scale_bar_artists = []
        self._colorbars = []
        self._molecule_artists = []
        # Reset profile artists as figure was cleared
        self._profile_line = None
        self._profile_p0 = None
        self._profile_p1 = None
        self._profile_echo_artists = []
        n = len(self.views)
        if n == 0:
            self.draw(); return
        if self._view_layout == "stacked":
            cols = 1
            rows = n
        else:
            cols = int(math.ceil(math.sqrt(n)))
            rows = int(math.ceil(n / cols))
        for i, v in enumerate(self.views):
            ax = self.fig.add_subplot(rows, cols, i+1)
            self._ax_view_map[ax] = v
            if i == 0:
                self.main_ax = ax
            arr = np.asarray(v['arr'])
            flip = bool(v.get('relative_axes'))
            if flip:
                arr_plot = np.flipud(arr)
            else:
                arr_plot = arr
            extent = v.get('extent', None)
            cmap = v.get('cmap', 'viridis')
            origin = 'lower' if flip else 'upper'
            if extent is None:
                im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
            else:
                im = ax.imshow(arr_plot, extent=extent, origin=origin, interpolation='nearest', aspect='equal', cmap=cmap)
            ax.set_autoscale_on(False)
            cbar_label = v.get('colorbar_label') or v.get('unit', '')
            if cbar_label and self._show_colorbar:
                try:
                    divider = make_axes_locatable(ax)
                    if self._colorbar_orientation == 'horizontal':
                        cax = divider.append_axes("bottom", size="5%", pad=0.08)
                        cbar = self.fig.colorbar(im, cax=cax, orientation='horizontal')
                        cbar.set_label(cbar_label)
                        cbar.ax.xaxis.set_label_coords(0.5, 0.5)
                        cbar.ax.xaxis.label.set_horizontalalignment('center')
                        cbar.ax.xaxis.label.set_verticalalignment('center')
                    else:
                        cax = divider.append_axes("right", size="4%", pad=0.02)
                        cbar = self.fig.colorbar(im, cax=cax, orientation='vertical')
                        cbar.set_label(cbar_label)
                        cbar.ax.yaxis.set_label_coords(0.5, 0.5)
                        cbar.ax.yaxis.label.set_horizontalalignment('center')
                        cbar.ax.yaxis.label.set_verticalalignment('center')
                except Exception:
                    cbar = self.fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02, orientation=self._colorbar_orientation)
                    cbar.set_label(cbar_label)
                if not self._show_ticks:
                    cbar.set_ticks([])
                self._colorbars.append(cbar)
            title = v.get('title', '')
            ax.set_title(title, fontsize=9)
            ax.tick_params(labelsize=8)
            if not self._show_ticks:
                ax.set_xticks([])
                ax.set_yticks([])
            if self.scale_bar_enabled:
                try:
                    self._add_scale_bar(ax, v)
                except Exception:
                    pass
            # Draw molecules on every view
            self._draw_molecules(ax)
        try: self.fig.tight_layout()
        except Exception: pass
        self._apply_view_theme()
        self._apply_view_font_scale()
        # if profile mode is enabled, (re)create artists on main ax
        if self.profile_enabled:
            self._ensure_profile_artists()
            self._emit_profile()
        self.draw()

    def _draw_molecules(self, ax):
        if not self.molecules:
            return

        for mol in self.molecules:
            coords = mol.get_transformed_coordinates()
            if len(coords) == 0:
                continue
            
            # Z-range for depth cueing
            z_vals = coords[:, 2]
            z_min = z_vals.min()
            z_range = z_vals.ptp()
            if z_range < 1e-6:
                z_range = 1.0

            lc = None
            sc = None
            # Draw Bonds
            if 'Bonds' in mol.display_mode and mol.bonds:
                lines = []
                colors = []
                linewidths = []
                for (i, j) in mol.bonds:
                    if i >= len(coords) or j >= len(coords): continue
                    p1 = coords[i]
                    p2 = coords[j]
                    lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                    
                    # Depth cueing for bonds
                    z_mid = (p1[2] + p2[2]) * 0.5
                    z_norm = (z_mid - z_min) / z_range
                    alpha = 0.4 + 0.6 * z_norm
                    colors.append((0.9, 0.9, 0.9, alpha)) # White/Grey bonds
                    linewidths.append(1.0 + 2.0 * z_norm)
                
                lc = LineCollection(lines, colors=colors, linewidths=linewidths, zorder=29)
                ax.add_collection(lc)
                lc.set_pickradius(5) # Help hit testing if needed later

            # Draw Atoms
            if 'Atoms' in mol.display_mode:
                # Sort atoms by Z for simple painter's algorithm
                order = np.argsort(z_vals)
                coords_sorted = coords[order]
                elements_sorted = [mol.elements[i] for i in order]
                
                x = coords_sorted[:, 0]
                y = coords_sorted[:, 1]
                z = coords_sorted[:, 2]
                
                z_norm = (z - z_min) / z_range
                sizes = 40 + 80 * z_norm
                
                base_colors = [get_cpk_color(e) for e in elements_sorted]
                rgba_colors = [matplotlib.colors.to_rgba(c) for c in base_colors]
                final_colors = []
                for i, (r, g, b, a) in enumerate(rgba_colors):
                    depth_alpha = 0.4 + 0.6 * z_norm[i]
                    final_colors.append((r, g, b, depth_alpha))
                
                sc = ax.scatter(x, y, s=sizes, c=final_colors, edgecolors='black', linewidths=0.5, zorder=30)

            self._molecule_artists.append({
                'mol': mol,
                'ax': ax,
                'scatter': sc,
                'lines': lc
            })

    def _update_molecule_artists(self):
        """Update positions of existing molecule artists without full redraw."""
        for entry in self._molecule_artists:
            mol = entry['mol']
            sc = entry['scatter']
            lc = entry['lines']
            
            coords = mol.get_transformed_coordinates()
            if len(coords) == 0: continue
            
            # Re-calculate Z sort and props (same logic as _draw_molecules)
            z_vals = coords[:, 2]
            z_min = z_vals.min()
            z_range = z_vals.ptp()
            if z_range < 1e-6: z_range = 1.0
            
            if lc:
                # Update bonds
                lines = []
                colors = []
                linewidths = []
                for (i, j) in mol.bonds:
                    if i >= len(coords) or j >= len(coords): continue
                    p1 = coords[i]; p2 = coords[j]
                    lines.append([(p1[0], p1[1]), (p2[0], p2[1])])
                    z_mid = (p1[2] + p2[2]) * 0.5
                    z_norm = (z_mid - z_min) / z_range
                    alpha = 0.4 + 0.6 * z_norm
                    colors.append((0.9, 0.9, 0.9, alpha))
                    linewidths.append(1.0 + 2.0 * z_norm)
                lc.set_segments(lines)
                lc.set_color(colors)
                lc.set_linewidths(linewidths)

            if sc:
                # Update atoms
                order = np.argsort(z_vals)
                coords_sorted = coords[order]
                elements_sorted = [mol.elements[i] for i in order]
                x = coords_sorted[:, 0]
                y = coords_sorted[:, 1]
                z_norm = (z_vals[order] - z_min) / z_range
                sc.set_offsets(np.c_[x, y])
                # Note: updating sizes/colors is possible but expensive; 
                # for pure translation we could skip it, but rotation needs it.
                # We'll skip full color/size recalc for speed during drag if needed, 
                # but for now let's do it to keep depth cues correct.
                sizes = 40 + 80 * z_norm
                base_colors = [get_cpk_color(e) for e in elements_sorted]
                rgba_colors = [matplotlib.colors.to_rgba(c) for c in base_colors]
                final_colors = []
                for i, (r, g, b, a) in enumerate(rgba_colors):
                    depth_alpha = 0.4 + 0.6 * z_norm[i]
                    final_colors.append((r, g, b, depth_alpha))
                sc.set_sizes(sizes)
                sc.set_facecolors(final_colors)
        
        self.draw_idle()

    def enable_scale_bar(self, enable: bool):
        if enable == self.scale_bar_enabled:
            return
        self.scale_bar_enabled = enable
        if enable:
            self._connect_scale_bar_events()
        else:
            self._disconnect_scale_bar_events()
        self._redraw()

    def _connect_scale_bar_events(self):
        if self._scale_bar_cids:
            return
        self._scale_bar_cids = [
            self.mpl_connect('button_press_event', self._on_sb_press),
            self.mpl_connect('motion_notify_event', self._on_sb_motion),
            self.mpl_connect('button_release_event', self._on_sb_release),
        ]

    def _disconnect_scale_bar_events(self):
        for cid in self._scale_bar_cids:
            self.mpl_disconnect(cid)
        self._scale_bar_cids = []

    def _calculate_best_scale_bar(self, width, unit):
        if width <= 0:
            return 1.0, unit
        # Target roughly 15-20% of the image width
        target = width * 0.18
        exponent = math.floor(math.log10(target))
        fraction = target / (10**exponent)
        
        # Candidates for "elegant" sizes: 1, 2, 3, 4, 5, 10
        candidates = [1, 2, 3, 4, 5, 10]
        best_mantissa = min(candidates, key=lambda x: abs(x - fraction))
        size = best_mantissa * (10**exponent)
        
        # Auto-format label for common units
        label = f"{size:g} {unit}"
        if unit == 'nm':
            if size < 1.0:
                label = f"{size*1000:.0f} pm"
            elif size >= 1000:
                label = f"{size/1000:.2g} µm"
            else:
                label = f"{size:g} nm"
        elif unit == 'µm':
            if size < 1.0:
                label = f"{size*1000:.0f} nm"
            else:
                label = f"{size:g} µm"
        
        return size, label

    def _add_scale_bar(self, ax, view):
        extent = view.get('extent')
        if extent is None:
            # Fallback for pixel coords
            h, w = np.shape(view['arr'])
            width = w
            unit = 'px'
        else:
            width = abs(extent[1] - extent[0])
            unit = view.get('axis_unit') or 'nm'
            
        size, label = self._calculate_best_scale_bar(width, unit)
        
        font_scale = getattr(self, '_view_font_scale', 1.0)
        sb = AnchoredSizeBar(ax.transData, size, label, 
                             loc='center',  # Anchor point on the artist itself
                             pad=0.4, borderpad=0, sep=3, 
                             frameon=False, 
                             size_vertical=width*0.004*font_scale,
                             color='white',
                             label_top=True,
                             bbox_to_anchor=self._scale_bar_pos,
                             bbox_transform=ax.transAxes)
        
        # Apply font scaling
        sb.size_bar.get_children()[0].set_linewidth(0) # remove border if any
        text = sb.txt_label.get_children()[0]
        
        text.set_fontsize(10 * font_scale)
        text.set_fontweight('bold')
        sb.set_zorder(20)
        
        ax.add_artist(sb)
        self._scale_bar_artists.append(sb)

    def _on_sb_press(self, event):
        if not self.scale_bar_enabled: return
        
        # Check if we clicked a scale bar
        target_sb = None
        for sb in self._scale_bar_artists:
            if sb.contains(event)[0]:
                target_sb = sb
                break
        
        if target_sb is None:
            return

        if event.button == 1:
            self._scale_bar_drag_start = (event.x, event.y)
        elif event.button == 3:
            self._show_sb_context_menu(event)

    def _show_sb_context_menu(self, event):
        menu = QtWidgets.QMenu(self)
        
        col_menu = menu.addMenu("Colors")
        txt_act = col_menu.addAction("Text Color")
        bar_act = col_menu.addAction("Bar Color")
        
        font_menu = menu.addMenu("Font")
        # Top common fonts in Python/World (Windows/Linux/Mac safe-ish subset)
        fonts = [
            "Arial", "DejaVu Sans", "Times New Roman", "Courier New",
            "Verdana", "Tahoma", "Georgia", "Segoe UI",
            "Trebuchet MS", "Impact", "Calibri", "Cambria"
        ]
        for font_name in fonts:
            act = font_menu.addAction(font_name)
            # Show font in its own style
            try:
                f = QtGui.QFont(font_name)
                f.setPointSize(10)
                act.setFont(f)
            except Exception:
                pass
            act.triggered.connect(lambda checked, f=font_name: self._set_sb_font(f))
            
        txt_act.triggered.connect(self._pick_sb_text_color)
        bar_act.triggered.connect(self._pick_sb_bar_color)
        
        if getattr(event, 'guiEvent', None):
            menu.exec_(event.guiEvent.globalPos())

    def _set_sb_font(self, font):
        self._scale_bar_settings['font_family'] = font
        self._redraw()

    def _pick_sb_text_color(self):
        col = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, self, "Select Text Color")
        if col.isValid():
            self._scale_bar_settings['text_color'] = col.name()
            self._redraw()

    def _pick_sb_bar_color(self):
        col = QtWidgets.QColorDialog.getColor(QtCore.Qt.white, self, "Select Bar Color")
        if col.isValid():
            self._scale_bar_settings['bar_color'] = col.name()
            self._redraw()

    def _on_sb_motion(self, event):
        if self._scale_bar_drag_start is None:
            if self.scale_bar_enabled and event.inaxes:
                for sb in self._scale_bar_artists:
                    if sb.contains(event)[0]:
                        self.setCursor(QtCore.Qt.SizeAllCursor)
                        return
            self.setCursor(QtCore.Qt.ArrowCursor)
            return

        if event.inaxes is None: return
        ax = event.inaxes
        dx = (event.x - self._scale_bar_drag_start[0]) / ax.bbox.width
        dy = (event.y - self._scale_bar_drag_start[1]) / ax.bbox.height
        cur_x, cur_y = self._scale_bar_pos
        self._scale_bar_pos = (cur_x + dx, cur_y + dy)
        self._scale_bar_drag_start = (event.x, event.y)
        
        for sb in self._scale_bar_artists:
            if sb.axes:
                sb.set_bbox_to_anchor(self._scale_bar_pos, sb.axes.transAxes)
            
        self.draw_idle()

    def _on_sb_release(self, event):
        self._scale_bar_drag_start = None

    # ---------- Interactive profile helpers ----------
    def set_profile_callback(self, cb):
        self.profile_callback = cb

    def set_profile_highlight_callback(self, cb):
        self._profile_highlight_cb = cb

    def set_profile_label_scale(self, scale):
        try:
            scale = float(scale)
        except Exception:
            return
        scale = max(0.6, min(2.5, scale))
        if abs(scale - self._profile_label_scale) <= 1e-3:
            return
        self._profile_label_scale = scale
        self._update_profile_markers()
        for entry in self._saved_profiles:
            text = entry.get('label_artist')
            base = entry.get('label_base_size', 8.0)
            if text is not None:
                try:
                    text.set_fontsize(base * self._profile_label_scale)
                except Exception:
                    pass
        self.draw_idle()

    def set_profile_marker_callback(self, cb):
        self._profile_marker_callback = cb

    def set_profile_state_callback(self, cb):
        self._profile_state_callback = cb

    def export_profile_state(self):
        saved = []
        for entry in self._saved_profiles:
            pts = entry.get('pts')
            if pts is None:
                continue
            saved.append({'pts': tuple(pts), 'color': entry.get('color'), 'lw': entry.get('lw')})
        state = {
            'active_pts': tuple(self.profile_pts) if self.profile_pts is not None else None,
            'saved': saved,
            'marker_key': self._profile_marker_key,
            'marker_positions_by_key': dict(self._profile_marker_positions_by_key),
            'marker_domain_by_key': dict(self._profile_marker_domain_by_key),
        }
        return state

    def export_profile_datasets(self):
        """Return active/saved profile datasets for external dialogs."""
        active = self._build_profile_data(self.profile_pts, color=self._active_profile_color)
        saved = []
        for entry in self._saved_profiles:
            data = entry.get('data')
            if data is None:
                data = self._build_profile_data(entry.get('pts'), color=entry.get('color'))
                entry['data'] = data
            if data:
                saved.append(data)
        return active, saved

    def import_profile_state(self, state, emit=True):
        if state is None:
            return
        if self._profile_state_syncing:
            return
        try:
            self._profile_state_syncing = True
            active_pts = state.get('active_pts')
            saved = state.get('saved') or []
            self._profile_marker_key = state.get('marker_key')
            self._profile_marker_positions_by_key = dict(state.get('marker_positions_by_key') or {})
            self._profile_marker_domain_by_key = dict(state.get('marker_domain_by_key') or {})
            if active_pts is not None:
                self._set_profile_pts(tuple(active_pts))
                if not self.profile_enabled:
                    self.enable_profile(True)
            self._clear_saved_profile_artists(notify=False)
            for entry in saved:
                pts = entry.get('pts')
                if pts is None:
                    continue
                self._add_saved_profile_from_pts(tuple(pts), entry.get('color'), entry.get('lw'))
            self._ensure_profile_artists()
            self._update_profile_artists()
            self.set_profile_marker_key(self._profile_marker_key)
        finally:
            self._profile_state_syncing = False
        if emit:
            self._emit_profile_state()

    def set_profile_marker_key(self, key):
        self._profile_marker_key = key
        if key is None:
            self._profile_marker_positions = self._profile_marker_positions_by_key.get(None)
            self._profile_marker_domain = self._profile_marker_domain_by_key.get(None)
        else:
            try:
                idx = int(key)
            except Exception:
                idx = None
            if idx is None:
                self._profile_marker_positions = None
                self._profile_marker_domain = None
            else:
                self._profile_marker_positions = self._profile_marker_positions_by_key.get(idx)
                self._profile_marker_domain = self._profile_marker_domain_by_key.get(idx)
        self._update_profile_marker_artists()
        self._update_profile_hud()

    def set_profile_marker_positions(self, positions, domain=None, emit=True, profile_key=None):
        key = self._profile_marker_key if profile_key is None else profile_key
        if positions is None or len(positions) < 2:
            self._profile_marker_positions = None
            if domain is not None:
                self._profile_marker_domain = tuple(domain)
            if key is not None:
                self._profile_marker_positions_by_key.pop(key, None)
                if domain is not None:
                    self._profile_marker_domain_by_key[key] = tuple(domain)
            self._clear_profile_marker_artists()
            if emit and callable(self._profile_marker_callback):
                self._profile_marker_callback(None, None)
            return
        if domain is not None:
            self._profile_marker_domain = tuple(domain)
        self._profile_marker_positions = [float(p) for p in positions]
        if key is not None:
            self._profile_marker_positions_by_key[key] = list(self._profile_marker_positions)
            if domain is not None:
                self._profile_marker_domain_by_key[key] = tuple(domain)
        self._update_profile_marker_artists()
        if emit and callable(self._profile_marker_callback):
            self._profile_marker_callback(list(self._profile_marker_positions),
                                          tuple(self._profile_marker_domain) if self._profile_marker_domain else None)

    def set_detail_theme(self, *, dark=None, grid=None):
        changed = False
        if dark is not None and bool(dark) != self._detail_dark:
            self._detail_dark = bool(dark)
            changed = True
        if grid is not None and bool(grid) != self._detail_grid:
            self._detail_grid = bool(grid)
            changed = True
        if changed:
            self._apply_view_theme()

    def set_angle_callback(self, cb):
        self.angle_callback = cb

    def enable_angle(self, enable: bool):
        if enable == self.angle_enabled:
            return
        self.angle_enabled = enable
        if enable:
            self._connect_angle_events()
            self._ensure_angle_artists()
            self._emit_angle()
        else:
            self._disconnect_angle_events()
            self._clear_angle_artists()
            self.angle_pts = None
            self._emit_angle()
        self.draw_idle()

    def clear_angle_measurement(self):
        self.angle_pts = None
        self._clear_angle_artists()
        if self.angle_enabled:
            self._ensure_angle_artists()
            self._emit_angle()

    def _connect_angle_events(self):
        if self._angle_cids:
            return
        self._angle_cids = [
            self.mpl_connect('button_press_event', self._on_angle_press),
            self.mpl_connect('button_release_event', self._on_angle_release),
            self.mpl_connect('motion_notify_event', self._on_angle_motion),
        ]

    def _disconnect_angle_events(self):
        for cid in self._angle_cids:
            try:
                self.mpl_disconnect(cid)
            except Exception:
                pass
        self._angle_cids = []

    def _apply_view_theme(self):
        dark = bool(self._detail_dark)
        fig_face = '#111217' if dark else '#ffffff'
        ax_face = '#14161c' if dark else '#ffffff'
        text_color = '#f5f5f5' if dark else '#111111'
        grid_color = '#4f5a64' if dark else '#9a9a9a'
        try:
            self.fig.set_facecolor(fig_face)
        except Exception:
            pass
        for ax in self.fig.axes:
            try:
                is_colorbar = ax in [cbar.ax for cbar in self._colorbars]
            except Exception:
                is_colorbar = False
            try:
                ax.set_facecolor(ax_face if not is_colorbar else fig_face)
                ax.tick_params(colors=text_color, labelcolor=text_color)
                ax.xaxis.label.set_color(text_color)
                ax.yaxis.label.set_color(text_color)
                for spine in ax.spines.values():
                    spine.set_color(text_color)
                if not is_colorbar:
                    if self._detail_grid:
                        ax.grid(True, color=grid_color, alpha=0.3, linewidth=0.6)
                    else:
                        ax.grid(False)
            except Exception:
                pass
        for cbar in getattr(self, '_colorbars', []):
            try:
                cbar.ax.tick_params(colors=text_color, labelcolor=text_color)
                cbar.ax.yaxis.label.set_color(text_color)
                cbar.ax.xaxis.label.set_color(text_color)
                cbar.outline.set_edgecolor(text_color)
            except Exception:
                pass
        # Update scale bar colors
        sb_settings = getattr(self, '_scale_bar_settings', {})
        sb_text_col = sb_settings.get('text_color') or text_color
        sb_bar_col = sb_settings.get('bar_color') or text_color
        
        for sb in self._scale_bar_artists:
            try:
                sb.size_bar.get_children()[0].set_color(sb_bar_col)
                sb.txt_label.get_children()[0].set_color(sb_text_col)
            except Exception:
                pass
        if self.angle_pts:
            self._update_angle_artists()
        self.draw_idle()

    def _apply_view_font_scale(self):
        scale = max(0.6, min(2.5, getattr(self, '_view_font_scale', 1.0)))
        tick_size = 8 * scale
        label_size = 10 * scale
        title_size = 9 * scale
        for ax in self.fig.axes:
            try:
                ax.tick_params(labelsize=tick_size)
                ax.xaxis.label.set_fontsize(label_size)
                ax.yaxis.label.set_fontsize(label_size)
                ax.title.set_fontsize(title_size)
            except Exception:
                pass
        for cbar in getattr(self, '_colorbars', []):
            try:
                cbar.ax.tick_params(labelsize=tick_size)
                cbar.ax.yaxis.label.set_fontsize(label_size)
                cbar.ax.xaxis.label.set_fontsize(label_size)
            except Exception:
                pass
        # Update scale bar font size
        for sb in self._scale_bar_artists:
            try:
                sb.txt_label.get_children()[0].set_fontsize(10 * scale)
            except Exception:
                pass
        self.draw_idle()
        if getattr(self, '_angle_label', None) is not None:
            try:
                self._angle_label.set_fontsize(9 * scale)
            except Exception:
                pass
        for lbl in getattr(self, '_angle_len_labels', []):
            try:
                lbl.set_fontsize(8 * scale)
            except Exception:
                pass

    def _emit_angle(self):
        if not callable(self.angle_callback):
            return
        info = self._compute_angle_info()
        try:
            self.angle_callback(info)
        except Exception:
            pass

    def set_copy_feedback_handler(self, handler):
        self._copy_feedback_handler = handler

    def get_overview_pixmap(self):
        """Return a pixmap snapshot of the current canvas (with overlays)."""
        try:
            return self.grab()
        except Exception:
            return None

    def set_value_callback(self, cb):
        self._value_callback = cb

    def wheelEvent(self, event):
        try:
            mods = event.modifiers()
        except Exception:
            mods = QtCore.Qt.NoModifier
        if mods & QtCore.Qt.ControlModifier:
            delta = event.angleDelta().y() if hasattr(event, 'angleDelta') else 0
            if delta:
                step = 0.05 * (1 if delta > 0 else -1)
                self._view_font_scale = min(2.5, max(0.6, self._view_font_scale + step))
                self._apply_view_font_scale()
            event.accept()
            return
        super().wheelEvent(event)

    def enable_profile(self, enable:bool):
        if enable == self.profile_enabled:
            return
        self.profile_enabled = enable
        if enable:
            self._connect_profile_events()
            self._ensure_profile_artists()
        else:
            self._disconnect_profile_events()
            self._clear_profile_artists()
            self.profile_pts = None
            self._clear_saved_profile_artists(notify=False)
            self._active_profile_original_color = None
            self._profile_marker_positions = None
            self._profile_marker_domain = None
            self._clear_profile_hud()
        self.draw_idle()

    def keyPressEvent(self, event):
        if self.profile_enabled and event is not None:
            try:
                if event.modifiers() & QtCore.Qt.ControlModifier and event.key() == QtCore.Qt.Key_Z:
                    self._undo_last_profile_snapshot()
                    return
            except Exception:
                pass
        super().keyPressEvent(event)

    def _connect_profile_events(self):
        if self._cids:
            return
        self._cids = [
            self.mpl_connect('button_press_event', self._on_press),
            self.mpl_connect('button_release_event', self._on_release),
            self.mpl_connect('motion_notify_event', self._on_motion),
        ]

    def _disconnect_profile_events(self):
        for cid in self._cids:
            try: self.mpl_disconnect(cid)
            except Exception: pass
        self._cids = []

    def _ensure_profile_artists(self):
        if self.main_ax is None:
            return
        if self._profile_line is None:
            # initialize points centered if not set
            if self.profile_pts is None:
                try:
                    v0 = self.views[0]
                    arr = np.asarray(v0['arr'])
                    h, w = arr.shape
                    if v0.get('extent') is None:
                        x0 = w*0.25; y0 = h*0.5; x1 = w*0.75; y1 = h*0.5
                    else:
                        xmin, xmax, ymin, ymax = v0['extent'][0], v0['extent'][1], v0['extent'][2], v0['extent'][3]
                        # our code sets extent as [0, XRange, YRange, 0], so choose a centered horizontal line
                        x0 = xmin + 0.25*(xmax - xmin); x1 = xmin + 0.75*(xmax - xmin)
                        y0 = ymax + 0.5*(ymin - ymax); y1 = y0
                except Exception:
                    x0 = 0.25; x1 = 0.75; y0 = y1 = 0.5
                self._set_profile_pts((x0, y0, x1, y1))
            x0, y0, x1, y1 = self.profile_pts or (x0, y0, x1, y1)
            self._set_profile_pts((x0, y0, x1, y1))
            x0, y0, x1, y1 = self.profile_pts
            color = self._active_profile_color
            self._profile_line, = self.main_ax.plot([x0,x1],[y0,y1], color=color, lw=self._active_profile_lw, alpha=0.95, zorder=9)
            self._profile_p0, = self.main_ax.plot([x0],[y0], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
            self._profile_p1, = self.main_ax.plot([x1],[y1], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
            self._profile_endpoint_labels = self._create_endpoint_labels((x0, y0, x1, y1), color)
            self._profile_label = self._create_profile_id_label((x0, y0, x1, y1), "Active", color)
            self._update_profile_markers()
        
        # Clear existing echo artists to prevent duplicates
        for entry in self._profile_echo_artists:
            for art in entry.values():
                try: art.remove()
                except Exception: pass
        self._profile_echo_artists = []
        x0, y0, x1, y1 = self.profile_pts
        color = self._active_profile_color
        for ax in self._ax_view_map:
            if ax is self.main_ax:
                continue
            try:
                l, = ax.plot([x0,x1],[y0,y1], color=color, lw=self._active_profile_lw, alpha=0.95, zorder=9)
                p0, = ax.plot([x0],[y0], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
                p1, = ax.plot([x1],[y1], marker='o', color=color, ms=7, mec='black', mew=1.0, zorder=10)
                self._profile_echo_artists.append({'line': l, 'p0': p0, 'p1': p1})
            except Exception:
                pass

    def _ensure_angle_artists(self):
        if self.main_ax is None:
            return
        if self.angle_pts is None:
            try:
                xmin, xmax, ymin, ymax = self._angle_bounds()
            except Exception:
                xmin = ymin = 0.0
                xmax = ymax = 1.0
            vx = (xmin + xmax) * 0.5
            vy = (ymin + ymax) * 0.5
            ax = vx + 0.2 * (xmax - xmin)
            ay = vy
            bx = vx
            by = vy + 0.2 * (ymax - ymin)
            self.angle_pts = (vx, vy, ax, ay, bx, by)
        vx, vy, ax, ay, bx, by = self.angle_pts
        if not self._angle_lines:
            color1 = '#ffb300'
            color2 = '#00acc1'
            line1, = self.main_ax.plot([vx, ax], [vy, ay], color=color1, lw=2.4, alpha=0.95, zorder=9)
            line2, = self.main_ax.plot([vx, bx], [vy, by], color=color2, lw=2.4, alpha=0.95, zorder=9)
            vertex, = self.main_ax.plot([vx], [vy], marker='o', color='#ffffff', mec='#000000', ms=7, zorder=10)
            end_a, = self.main_ax.plot([ax], [ay], marker='o', color=color1, mec='#000000', ms=6, zorder=10)
            end_b, = self.main_ax.plot([bx], [by], marker='o', color=color2, mec='#000000', ms=6, zorder=10)
            self._angle_lines = [line1, line2]
            self._angle_markers = [vertex, end_a, end_b]
        self._update_angle_artists()

    def _clear_angle_artists(self):
        for art in self._angle_lines + self._angle_markers:
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._angle_lines = []
        self._angle_markers = []
        if self._angle_label is not None:
            try:
                self._angle_label.remove()
            except Exception:
                pass
        self._angle_label = None
        for lbl in self._angle_len_labels:
            try:
                lbl.remove()
            except Exception:
                pass
        self._angle_len_labels = []
        if self._angle_patch is not None:
            try:
                self._angle_patch.remove()
            except Exception:
                pass
        self._angle_patch = None

    def _update_angle_artists(self):
        if not self.angle_pts or not self._angle_lines:
            return
        vx, vy, ax, ay, bx, by = self.angle_pts
        self._angle_lines[0].set_data([vx, ax], [vy, ay])
        self._angle_lines[1].set_data([vx, bx], [vy, by])
        self._angle_markers[0].set_data([vx], [vy])
        self._angle_markers[1].set_data([ax], [ay])
        self._angle_markers[2].set_data([bx], [by])
        self._update_angle_label()
        self.draw_idle()

    def _update_angle_label(self):
        angle_info = self._compute_angle_info()
        if self._angle_label is not None:
            try:
                self._angle_label.remove()
            except Exception:
                pass
            self._angle_label = None
        for lbl in self._angle_len_labels:
            try:
                lbl.remove()
            except Exception:
                pass
        self._angle_len_labels = []
        if self._angle_patch is not None:
            try:
                self._angle_patch.remove()
            except Exception:
                pass
            self._angle_patch = None
        if not angle_info:
            self.draw_idle()
            return
        vx, vy = self.angle_pts[0], self.angle_pts[1]
        text = f"{angle_info['angle_deg']:.1f}-"
        unit = angle_info.get('unit')
        color = '#f5f5f5' if self._detail_dark else '#111111'
        bbox_face = '#060606' if self._detail_dark else 'white'
        font_scale = getattr(self, '_view_font_scale', 1.0)
        self._angle_label = self.main_ax.text(
            vx, vy, text,
            color=color,
            fontsize=9 * font_scale,
            ha='center', va='center',
            bbox={'facecolor': bbox_face, 'alpha': 0.65 if self._detail_dark else 0.7, 'edgecolor': 'none', 'pad': 2},
            zorder=12)
        # position label along bisector
        vec_a = np.array([self.angle_pts[2] - vx, self.angle_pts[3] - vy], dtype=float)
        vec_b = np.array([self.angle_pts[4] - vx, self.angle_pts[5] - vy], dtype=float)
        len_a = angle_info['len_a'] or 1.0
        len_b = angle_info['len_b'] or 1.0
        bis = vec_a / max(len_a, 1e-9) + vec_b / max(len_b, 1e-9)
        if np.allclose(bis, 0):
            bis = np.array([-(vec_a[1]), vec_a[0]])
        bis = bis / (np.linalg.norm(bis) + 1e-9)
        offset = min(len_a, len_b) * 0.2
        bx = vx + bis[0] * offset
        by = vy + bis[1] * offset
        self._angle_label.set_position((bx, by))
        # draw wedge patch
        theta_a = math.degrees(math.atan2(vec_a[1], vec_a[0]))
        theta_b = math.degrees(math.atan2(vec_b[1], vec_b[0]))
        theta1, theta2 = theta_a, theta_b
        diff = (theta2 - theta1) % 360.0
        if diff > 180:
            theta1, theta2 = theta2, theta1
        radius = min(len_a, len_b) * 0.25
        radius = max(radius, 1e-3)
        wedge = patches.Wedge((vx, vy), radius, theta1, theta2,
                              facecolor=color, alpha=0.15, edgecolor='none', zorder=8)
        self._angle_patch = wedge
        self.main_ax.add_patch(wedge)
        # length labels along arms
        if unit:
            mid_a = (vx + (vec_a[0] * 0.6), vy + (vec_a[1] * 0.6))
            mid_b = (vx + (vec_b[0] * 0.6), vy + (vec_b[1] * 0.6))
            lbl_a = self.main_ax.text(mid_a[0], mid_a[1],
                                      f"{len_a:.2f} {unit}",
                                      color=color,
                                      fontsize=8 * font_scale,
                                      ha='center', va='bottom',
                                      bbox={'facecolor': bbox_face, 'alpha': 0.5 if self._detail_dark else 0.6,
                                            'edgecolor': 'none', 'pad': 1},
                                      zorder=12)
            lbl_b = self.main_ax.text(mid_b[0], mid_b[1],
                                      f"{len_b:.2f} {unit}",
                                      color=color,
                                      fontsize=8 * font_scale,
                                      ha='center', va='bottom',
                                      bbox={'facecolor': bbox_face, 'alpha': 0.5 if self._detail_dark else 0.6,
                                            'edgecolor': 'none', 'pad': 1},
                                      zorder=12)
            self._angle_len_labels = [lbl_a, lbl_b]

    def _compute_angle_info(self):
        if not self.angle_pts:
            return None
        vx, vy, ax, ay, bx, by = self.angle_pts
        vec_a = np.array([ax - vx, ay - vy], dtype=float)
        vec_b = np.array([bx - vx, by - vy], dtype=float)
        len_a = float(np.hypot(vec_a[0], vec_a[1]))
        len_b = float(np.hypot(vec_b[0], vec_b[1]))
        if len_a < 1e-9 or len_b < 1e-9:
            angle_deg = 0.0
        else:
            cosang = float(np.clip(np.dot(vec_a, vec_b) / (len_a * len_b), -1.0, 1.0))
            angle_deg = float(np.degrees(np.arccos(cosang)))
        unit = self._profile_axis_unit()
        return {'angle_deg': angle_deg, 'len_a': len_a, 'len_b': len_b, 'unit': unit, 'vertex': (vx, vy)}

    def _clear_profile_artists(self):
        for art in (self._profile_line, self._profile_p0, self._profile_p1):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_line = self._profile_p0 = self._profile_p1 = None
        self._remove_profile_markers()
        self._clear_profile_marker_artists()
        if self._profile_label is not None:
            try:
                self._profile_label.remove()
            except Exception:
                pass
        self._profile_label = None
        for lbl in self._profile_endpoint_labels:
            try:
                lbl.remove()
            except Exception:
                pass
        self._profile_endpoint_labels = []
        for entry in self._profile_echo_artists:
            for art in entry.values():
                try: art.remove()
                except Exception: pass
        self._profile_echo_artists = []
        self._clear_profile_hud()
        self.draw_idle()

    def _update_profile_artists(self):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0,x1],[y0,y1])
        self._profile_p0.set_data([x0],[y0])
        self._profile_p1.set_data([x1],[y1])
        for entry in self._profile_echo_artists:
            try:
                entry['line'].set_data([x0,x1],[y0,y1])
                entry['p0'].set_data([x0],[y0])
                entry['p1'].set_data([x1],[y1])
            except Exception: pass
        self._update_profile_markers()
        self._update_profile_marker_artists()
        self.draw_idle()
        self._emit_profile()

    def _update_profile_artists_fast(self, draw=True):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        if self.profile_pts is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0, x1], [y0, y1])
        self._profile_p0.set_data([x0], [y0])
        self._profile_p1.set_data([x1], [y1])
        self._update_profile_labels()
        if draw:
            self.draw_idle()

    def _schedule_profile_update(self):
        if not self._profile_update_timer.isActive():
            self._profile_update_timer.start()

    def _flush_profile_updates(self):
        if not self.profile_enabled:
            return
        if self.profile_pts is None:
            return
        self._update_profile_markers()
        self._emit_profile()
        self.draw_idle()

    def activate_saved_profile(self, index):
        """Promote a saved overlay back to the active profile line."""
        if index is None or not self._saved_profiles:
            return False
        try:
            idx = int(index)
        except Exception:
            return False
        if idx < 0 or idx >= len(self._saved_profiles):
            return False
        entry = self._saved_profiles.pop(idx)
        self._active_profile_original_color = entry.get('color')
        if self._profile_marker_positions_by_key:
            new_map = {}
            new_domain = {}
            moved_active = False
            for key, value in self._profile_marker_positions_by_key.items():
                if key is None:
                    new_map[key] = value
                    continue
                if key == idx:
                    new_map[None] = value
                    moved_active = True
                    continue
                if key > idx:
                    new_map[key - 1] = value
                else:
                    new_map[key] = value
            for key, value in self._profile_marker_domain_by_key.items():
                if key is None:
                    new_domain[key] = value
                    continue
                if key == idx:
                    new_domain[None] = value
                    continue
                if key > idx:
                    new_domain[key - 1] = value
                else:
                    new_domain[key] = value
            self._profile_marker_positions_by_key = new_map
            self._profile_marker_domain_by_key = new_domain
            if moved_active:
                self._profile_marker_key = None
                self._profile_marker_positions = new_map.get(None)
                self._profile_marker_domain = new_domain.get(None)
        # remove overlay artists from canvas
        for art in entry.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        # promote saved path to active profile
        pts = entry.get('pts')
        if pts is None:
            return False
        self._set_profile_pts(tuple(pts))
        self._ensure_profile_artists()
        self._update_profile_artists()
        self.draw_idle()
        self._emit_profile()
        return True

    def remove_saved_profile(self, index):
        """Remove a saved profile overlay by index."""
        if index is None:
            return False
        try:
            idx = int(index)
        except Exception:
            return False
        if idx < 0 or idx >= len(self._saved_profiles):
            return False
        self._remove_saved_profile(idx)
        return True

    def snapshot_active_profile(self):
        """Public hook: save the current active profile as an overlay."""
        self._snapshot_active_profile()

    def _remove_profile_markers(self):
        if getattr(self, '_profile_ticks', None) is not None:
            try:
                self._profile_ticks.remove()
            except Exception:
                pass
            self._profile_ticks = None
            
        if self._profile_info_text is not None:
            try:
                self._profile_info_text.remove()
            except Exception:
                pass
        self._profile_info_text = None

    def _clear_profile_hud(self):
        if self._profile_hud_text is not None:
            try:
                self._profile_hud_text.remove()
            except Exception:
                pass
        self._profile_hud_text = None

    def _create_profile_id_label(self, pts, text, color):
        if self.main_ax is None or pts is None:
            return None
        x0, y0, x1, y1 = pts
        xm = x0 + 0.5 * (x1 - x0)
        ym = y0 + 0.5 * (y1 - y0)
        try:
            return self.main_ax.text(
                xm, ym, text, color=color, fontsize=8,
                ha='center', va='bottom',
                bbox={'facecolor': 'black', 'alpha': 0.25, 'edgecolor': 'none', 'pad': 1.5},
                zorder=11)
        except Exception:
            return None

    def _create_endpoint_labels(self, pts, color):
        if self.main_ax is None or pts is None:
            return []
        x0, y0, x1, y1 = pts
        labels = []
        try:
            labels.append(self.main_ax.text(
                x0, y0, "A", color=color, fontsize=8,
                ha='right', va='bottom',
                bbox={'facecolor': 'black', 'alpha': 0.25, 'edgecolor': 'none', 'pad': 1.0},
                zorder=11))
            labels.append(self.main_ax.text(
                x1, y1, "B", color=color, fontsize=8,
                ha='left', va='bottom',
                bbox={'facecolor': 'black', 'alpha': 0.25, 'edgecolor': 'none', 'pad': 1.0},
                zorder=11))
        except Exception:
            return []
        return labels

    def _profile_axis_unit(self):
        if not self.views:
            return 'px'
        v0 = self.views[0]
        axis_unit = v0.get('axis_unit')
        if axis_unit:
            return axis_unit
        return 'px' if v0.get('extent') is None else 'nm'

    def _format_profile_label(self, pts):
        x0, y0, x1, y1 = pts
        dx = abs(float(x1) - float(x0))
        dy = abs(float(y1) - float(y0))
        length = float(math.hypot(dx, dy))
        unit = self._profile_axis_unit()
        return f"L={length:.3g} {unit} | dx={dx:.3g} {unit} | dy={dy:.3g} {unit}"

    def _create_ticks_and_label(self, pts, color, alpha=0.85, base_size=9):
        size = base_size * getattr(self, '_profile_label_scale', 1.0)
        try:
            fractions = (0.25, 0.5, 0.75)
            tx, ty = [], []
            for frac in fractions:
                x = pts[0] + (pts[2] - pts[0]) * frac
                y = pts[1] + (pts[3] - pts[1]) * frac
                tx.append(x)
                ty.append(y)
            ticks, = self.main_ax.plot(
                tx, ty, marker='s', linestyle='None', color=color,
                ms=max(3.0, 4.0 * self._profile_label_scale),
                alpha=alpha, zorder=9)
            label_text = self._format_profile_label(pts)
            xm = pts[0] + (pts[2] - pts[0]) * 0.5
            ym = pts[1] + (pts[3] - pts[1]) * 0.5
            text = self.main_ax.text(
                xm, ym, label_text, color=color, fontsize=size,
                ha='center', va='center',
                bbox={'facecolor': 'black', 'alpha': 0.35, 'edgecolor': 'none', 'pad': 2},
                zorder=11)
        except Exception:
            return None, None
        return ticks, text

    def _update_profile_markers(self):
        if self.profile_pts is None or self.main_ax is None:
            self._remove_profile_markers()
            return
        
        # Reuse existing artists if possible
        if self._profile_ticks is not None and self._profile_info_text is not None:
            self._remove_profile_markers() # Fallback to recreate if complex update needed, or optimize further
            ticks, text = self._create_ticks_and_label(self.profile_pts, color='yellow', alpha=0.9, base_size=9)
        else:
            self._remove_profile_markers()
            ticks, text = self._create_ticks_and_label(self.profile_pts, color='yellow', alpha=0.9, base_size=9)
            
        self._profile_ticks = ticks
        self._profile_info_text = text
        self._update_profile_marker_artists()
        self._update_profile_labels()
        self._update_profile_hud()

    def _update_profile_labels(self):
        if self.profile_pts is None or self.main_ax is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        if self._profile_endpoint_labels:
            try:
                self._profile_endpoint_labels[0].set_position((x0, y0))
                self._profile_endpoint_labels[1].set_position((x1, y1))
            except Exception:
                pass
        if self._profile_label is not None:
            try:
                xm = x0 + 0.5 * (x1 - x0)
                ym = y0 + 0.5 * (y1 - y0)
                self._profile_label.set_position((xm, ym))
            except Exception:
                pass

    def _update_profile_hud(self):
        if self.main_ax is None or self.profile_pts is None:
            self._clear_profile_hud()
            return
        key = self._profile_marker_key
        pts = self._profile_marker_pts() or self.profile_pts
        data = self._build_profile_data(pts, color=self._active_profile_color)
        if not data:
            self._clear_profile_hud()
            return
        unit = data.get('axis_unit') or data.get('distance_unit') or 'px'
        length = data.get('length_nm')
        title = "Active" if key is None else f"Overlay {int(key) + 1}"
        marker_delta = None
        if self._profile_marker_positions and self._profile_marker_domain:
            marker_delta = abs(self._profile_marker_positions[1] - self._profile_marker_positions[0])
        parts = [title]
        if length is not None:
            parts.append(f"L={length:.3g} {unit}")
        if marker_delta is not None:
            parts.append(f"?={marker_delta:.3g} {unit}")
        text = " | ".join(parts)
        if self._profile_hud_text is None:
            self._profile_hud_text = self.main_ax.text(
                0.02, 0.98, text, transform=self.main_ax.transAxes,
                ha='left', va='top', fontsize=9,
                color="#f5f5f5",
                bbox={'facecolor': 'black', 'alpha': 0.35, 'edgecolor': 'none', 'pad': 2},
                zorder=20)
        else:
            try:
                self._profile_hud_text.set_text(text)
            except Exception:
                pass

    def _clear_profile_marker_artists(self):
        for art in self._profile_marker_artists:
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_marker_artists = []
        self.draw_idle()

    def _profile_marker_points(self):
        pts = self._profile_marker_pts()
        if pts is None or self._profile_marker_positions is None:
            return []
        if not self._profile_marker_domain:
            return []
        x0, y0, x1, y1 = pts
        dom_min, dom_max = self._profile_marker_domain
        span = float(dom_max - dom_min) if dom_max != dom_min else 0.0
        if span == 0.0:
            return []
        points = []
        for pos in self._profile_marker_positions:
            t = (float(pos) - dom_min) / span
            t = max(0.0, min(1.0, t))
            px = x0 + (x1 - x0) * t
            py = y0 + (y1 - y0) * t
            points.append((px, py))
        return points

    def _profile_marker_pts(self):
        if self._profile_marker_key is None:
            return self.profile_pts
        try:
            idx = int(self._profile_marker_key)
        except Exception:
            return self.profile_pts
        if idx < 0 or idx >= len(self._saved_profiles):
            return self.profile_pts
        entry = self._saved_profiles[idx]
        return entry.get('pts') or self.profile_pts

    def _update_profile_marker_artists(self):
        pts = self._profile_marker_pts()
        if self.main_ax is None or pts is None:
            self._clear_profile_marker_artists()
            return
        points = self._profile_marker_points()
        if len(points) < 2:
            self._clear_profile_marker_artists()
            return
        for art in self._profile_marker_artists:
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_marker_artists = []
        color = '#ff5252'
        x0, y0, x1, y1 = pts
        vx = x1 - x0
        vy = y1 - y0
        length = float(math.hypot(vx, vy)) if vx or vy else 0.0
        if length > 0:
            nx = -vy / length
            ny = vx / length
        else:
            nx, ny = 0.0, 0.0
        tick_len = 0.03 * length if length > 0 else 0.0
        for px, py in points:
            if tick_len > 0:
                tick, = self.main_ax.plot(
                    [px - nx * tick_len, px + nx * tick_len],
                    [py - ny * tick_len, py + ny * tick_len],
                    color=color, lw=2.0, alpha=0.9, zorder=12,
                )
                self._profile_marker_artists.append(tick)
            marker, = self.main_ax.plot([px], [py], marker='o', color=color,
                                        ms=5, mec='white', mew=0.7, zorder=13)
            self._profile_marker_artists.append(marker)
        self.draw_idle()
        self._update_profile_hud()

    def _update_profile_marker_artists_fast(self):
        pts = self._profile_marker_pts()
        if self.main_ax is None or pts is None:
            return
        points = self._profile_marker_points()
        if len(points) < 2:
            return
        x0, y0, x1, y1 = pts
        vx = x1 - x0
        vy = y1 - y0
        length = float(math.hypot(vx, vy)) if vx or vy else 0.0
        if length > 0:
            nx = -vy / length
            ny = vx / length
        else:
            nx, ny = 0.0, 0.0
        tick_len = 0.03 * length if length > 0 else 0.0
        expected = len(points) * (2 if tick_len > 0 else 1)
        if len(self._profile_marker_artists) != expected:
            self._update_profile_marker_artists()
            return
        idx = 0
        for px, py in points:
            if tick_len > 0:
                try:
                    self._profile_marker_artists[idx].set_data(
                        [px - nx * tick_len, px + nx * tick_len],
                        [py - ny * tick_len, py + ny * tick_len],
                    )
                except Exception:
                    pass
                idx += 1
            try:
                self._profile_marker_artists[idx].set_data([px], [py])
            except Exception:
                pass
            idx += 1
        self.draw_idle()

    def _profile_marker_hit(self, x, y):
        points = self._profile_marker_points()
        if not points:
            return None
        min_idx = None
        min_dist = float('inf')
        for idx, (px, py) in enumerate(points):
            dist = self._pt_distance_pixels(x, y, px, py)
            if dist < min_dist:
                min_dist = dist
                min_idx = idx
        if min_dist <= 12.0:
            return min_idx
        return None

    def _build_profile_data(self, pts, color=None, view=None):
        if pts is None or not self.views:
            return None
        try:
            v0 = view if view is not None else self.views[0]
            arr = np.asarray(v0['arr'], dtype=float)
            h, w = arr.shape
            extent = v0.get('extent', None)
            axis_unit = self._profile_axis_unit()
            x0, y0, x1, y1 = pts
            if extent is None:
                c0 = x0; r0 = y0; c1 = x1; r1 = y1
                length_nm = None
            else:
                xmin, xmax = extent[0], extent[1]
                ymin, ymax = extent[2], extent[3]
                xr = (xmax - xmin) if (xmax is not None and xmin is not None) else 1.0
                yr = (ymin - ymax) if (ymin is not None and ymax is not None) else 1.0
                c0 = (x0 - xmin) / (xr + 1e-12) * (w - 1)
                c1 = (x1 - xmin) / (xr + 1e-12) * (w - 1)
                r0 = (y0 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                r1 = (y1 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                try:
                    dx_nm = (x1 - x0); dy_nm = (y1 - y0)
                    length_nm = float(math.hypot(dx_nm, dy_nm))
                except Exception:
                    length_nm = None
            n = int(max(2, round(((c1 - c0)**2 + (r1 - r0)**2) ** 0.5) + 1))
            t = np.linspace(0.0, 1.0, n)
            cc = c0 + (c1 - c0) * t
            rr = r0 + (r1 - r0) * t
            rr = np.clip(rr, 0, h - 1)
            cc = np.clip(cc, 0, w - 1)
            i0 = np.floor(rr).astype(int)
            j0 = np.floor(cc).astype(int)
            i1 = np.clip(i0 + 1, 0, h - 1)
            j1 = np.clip(j0 + 1, 0, w - 1)
            wy = rr - i0
            wx = cc - j0
            vals = (
                (1 - wy) * (1 - wx) * arr[i0, j0] +
                wy * (1 - wx) * arr[i1, j0] +
                (1 - wy) * wx * arr[i0, j1] +
                wy * wx * arr[i1, j1]
            )
            x_px = np.linspace(0.0, float(n - 1), n)
            unit = v0.get('unit', None)
            x_phys = None
            distance_unit = 'px'
            if length_nm is not None and n > 1:
                try:
                    scale = float(length_nm) / float(n - 1)
                    x_phys = x_px * scale
                    if axis_unit:
                        distance_unit = axis_unit
                except Exception:
                    x_phys = None
            meta = v0.get('meta') if isinstance(v0, dict) else None
            return {
                'x_px': x_px,
                'x_nm': x_phys,
                'vals': vals,
                'length_nm': length_nm,
                'unit': unit,
                'axis_unit': axis_unit,
                'distance_unit': distance_unit if x_phys is not None else 'px',
                'color': color,
                'label': self._format_profile_label(pts),
                'relative_axes': bool(v0.get('relative_axes')),
                'meta': meta,
            }
        except Exception:
            return None

    def _pt_distance_pixels(self, x, y, xp, yp):
        try:
            p_scr = self.main_ax.transData.transform((x, y))
            q_scr = self.main_ax.transData.transform((xp, yp))
            dx = p_scr[0] - q_scr[0]; dy = p_scr[1] - q_scr[1]
            return (dx*dx + dy*dy) ** 0.5
        except Exception:
            return float('inf')

    def _profile_bounds(self):
        try:
            v0 = self.views[0]
            extent = v0.get('extent')
            arr = np.asarray(v0['arr'])
            h, w = arr.shape
            if extent is None:
                return (0.0, float(w - 1), 0.0, float(h - 1))
            x0, x1, y1, y0 = extent
            return (min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1))
        except Exception:
            return (-1e6, 1e6, -1e6, 1e6)

    def _clamp_profile_pts(self, x0, y0, x1, y1):
        xmin, xmax, ymin, ymax = self._profile_bounds()
        return (
            max(xmin, min(xmax, x0)),
            max(ymin, min(ymax, y0)),
            max(xmin, min(xmax, x1)),
            max(ymin, min(ymax, y1)),
        )

    def _set_profile_pts(self, pts):
        if pts is None:
            self.profile_pts = None
            return
        x0, y0, x1, y1 = pts
        self.profile_pts = self._clamp_profile_pts(x0, y0, x1, y1)

    def _angle_bounds(self):
        return self._profile_bounds()

    def _set_angle_pts(self, vx, vy, ax, ay, bx, by):
        xmin, xmax, ymin, ymax = self._angle_bounds()
        def clamp(val, lo, hi):
            return max(lo, min(hi, val))
        vx = clamp(vx, xmin, xmax); vy = clamp(vy, ymin, ymax)
        ax = clamp(ax, xmin, xmax); ay = clamp(ay, ymin, ymax)
        bx = clamp(bx, xmin, xmax); by = clamp(by, ymin, ymax)
        self.angle_pts = (vx, vy, ax, ay, bx, by)

    def _shift_pressed(self, event):
        key = getattr(event, 'key', None)
        if key and 'shift' in str(key).lower():
            return True
        gui = getattr(event, 'guiEvent', None)
        try:
            if gui is not None and gui.modifiers() & QtCore.Qt.ShiftModifier:
                return True
        except Exception:
            pass
        return False

    def _snapshot_active_profile(self):
        if self.profile_pts is None or self.main_ax is None:
            return
        pts = tuple(self.profile_pts)
        if self._active_profile_original_color:
            color = self._active_profile_original_color
            self._active_profile_original_color = None
        else:
            color = next(self._profile_color_cycle)
        lw = self._active_profile_lw
        line, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], color=color, lw=lw, alpha=0.7, zorder=6, linestyle='--')
        # Combine endpoints into one artist
        endpoints, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], marker='o', linestyle='None', color=color, ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        base_size = 8
        ticks, text = self._create_ticks_and_label(pts, color=color, alpha=0.7, base_size=base_size)
        overlay_idx = len(self._saved_profiles) + 1
        overlay_label = self._create_profile_id_label(pts, f"Overlay {overlay_idx}", color)
        if overlay_label is not None:
            try:
                overlay_label.set_visible(False)
            except Exception:
                pass
        endpoint_labels = self._create_endpoint_labels(pts, color)
        for lbl in endpoint_labels:
            try:
                lbl.set_visible(False)
            except Exception:
                pass
        artists = [line, endpoints]
        if ticks: artists.append(ticks)
        if text: artists.append(text)
        if overlay_label is not None:
            artists.append(overlay_label)
        artists += endpoint_labels
        data = self._build_profile_data(pts, color=color)
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels, 'lw': lw}
        if text is not None:
            entry['label_artist'] = text
            entry['label_base_size'] = base_size
        self._saved_profiles.append(entry)
        self._refresh_overlay_labels()
        self.draw_idle()
        self._emit_profile()

    def _distance_to_segment_pixels(self, x, y, pts):
        try:
            px, py = self.main_ax.transData.transform((x, y))
            x0, y0 = self.main_ax.transData.transform((pts[0], pts[1]))
            x1, y1 = self.main_ax.transData.transform((pts[2], pts[3]))
            vx, vy = x1 - x0, y1 - y0
            if vx == 0 and vy == 0:
                return ((px - x0)**2 + (py - y0)**2) ** 0.5
            t = ((px - x0) * vx + (py - y0) * vy) / (vx * vx + vy * vy)
            t = max(0.0, min(1.0, t))
            proj_x = x0 + t * vx
            proj_y = y0 + t * vy
            return ((px - proj_x)**2 + (py - proj_y)**2) ** 0.5
        except Exception:
            return float('inf')

    def _delete_snapshot_near(self, x, y):
        if x is None or y is None or not self._saved_profiles:
            return
        target = None
        for entry in reversed(self._saved_profiles):
            pts = entry.get('pts')
            if pts is None:
                continue
            dist = self._distance_to_segment_pixels(x, y, pts)
            if dist <= 12.0:
                target = entry
                break
        if target is None:
            return
        for art in target.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._saved_profiles.remove(target)
        self.draw_idle()
        self._emit_profile()

    def _remove_saved_profile(self, idx):
        if idx < 0 or idx >= len(self._saved_profiles):
            return
        entry = self._saved_profiles.pop(idx)
        if self._profile_marker_positions_by_key:
            new_map = {}
            new_domain = {}
            for key, value in self._profile_marker_positions_by_key.items():
                if key is None:
                    new_map[key] = value
                    continue
                if key == idx:
                    continue
                if key > idx:
                    new_map[key - 1] = value
                else:
                    new_map[key] = value
            for key, value in self._profile_marker_domain_by_key.items():
                if key is None:
                    new_domain[key] = value
                    continue
                if key == idx:
                    continue
                if key > idx:
                    new_domain[key - 1] = value
                else:
                    new_domain[key] = value
            self._profile_marker_positions_by_key = new_map
            self._profile_marker_domain_by_key = new_domain
            if self._profile_marker_key is not None:
                if self._profile_marker_key == idx:
                    self._profile_marker_key = None
                elif self._profile_marker_key > idx:
                    self._profile_marker_key -= 1
        for art in entry.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self.highlight_saved_profile(None)
        if self._profile_highlight_cb:
            try:
                self._profile_highlight_cb(None)
            except Exception:
                pass
        self._refresh_overlay_labels()
        self.draw_idle()
        self._emit_profile()

    def _overlay_index_near(self, x, y, thresh=10.0):
        if x is None or y is None or not self._saved_profiles:
            return None
        screen_pt = self.main_ax.transData.transform((x, y))
        for idx in reversed(range(len(self._saved_profiles))):
            entry = self._saved_profiles[idx]
            pts = entry.get('pts')
            if pts is None:
                continue
            try:
                x0, y0, x1, y1 = pts
                p0 = self.main_ax.transData.transform((x0, y0))
                p1 = self.main_ax.transData.transform((x1, y1))
                vx, vy = p1[0]-p0[0], p1[1]-p0[1]
                if vx == 0 and vy == 0:
                    dist = ((screen_pt[0]-p0[0])**2 + (screen_pt[1]-p0[1])**2) ** 0.5
                else:
                    t = ((screen_pt[0]-p0[0])*vx + (screen_pt[1]-p0[1])*vy)/(vx*vx+vy*vy)
                    t = max(0.0, min(1.0, t))
                    proj = (p0[0]+t*vx, p0[1]+t*vy)
                    dist = ((screen_pt[0]-proj[0])**2 + (screen_pt[1]-proj[1])**2) ** 0.5
                if dist <= thresh:
                    return idx
            except Exception:
                continue
        return None

    def _undo_last_profile_snapshot(self):
        if not self._saved_profiles:
            return
        entry = self._saved_profiles.pop()
        for art in entry.get('artists', []):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self.draw_idle()
        self._emit_profile()

    def _clear_saved_profile_artists(self, notify=False):
        for entry in self._saved_profiles:
            for art in entry.get('artists', []):
                try:
                    if art is not None:
                        art.remove()
                except Exception:
                    pass
        self._saved_profiles = []
        self._highlighted_overlay = None
        self.draw_idle()
        if notify:
            self._emit_profile()
        self._refresh_overlay_labels()

    def _add_saved_profile_from_pts(self, pts, color, lw=1.5):
        if pts is None or self.main_ax is None:
            return
        pts = tuple(pts)
        color = color or next(self._profile_color_cycle)
        lw = float(lw or 1.5)
        line, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]],
                                  color=color, lw=lw, alpha=0.7, zorder=6, linestyle='--')
        endpoints, = self.main_ax.plot([pts[0], pts[2]], [pts[1], pts[3]], marker='o', linestyle='None', color=color,
                                       ms=5, mec='black', mew=0.7, alpha=0.9, zorder=7)
        base_size = 8
        ticks, text = self._create_ticks_and_label(pts, color=color, alpha=0.7, base_size=base_size)
        overlay_idx = len(self._saved_profiles) + 1
        overlay_label = self._create_profile_id_label(pts, f"Overlay {overlay_idx}", color)
        if overlay_label is not None:
            try:
                overlay_label.set_visible(False)
            except Exception:
                pass
        endpoint_labels = self._create_endpoint_labels(pts, color)
        for lbl in endpoint_labels:
            try:
                lbl.set_visible(False)
            except Exception:
                pass
        artists = [line, endpoints]
        if ticks: artists.append(ticks)
        if text: artists.append(text)
        if overlay_label is not None:
            artists.append(overlay_label)
        artists += endpoint_labels
        data = self._build_profile_data(pts, color=color)
        entry = {'artists': artists, 'pts': pts, 'color': color, 'data': data,
                 'overlay_label_artist': overlay_label, 'endpoint_labels': endpoint_labels, 'lw': lw}
        if text is not None:
            entry['label_artist'] = text
            entry['label_base_size'] = base_size
        self._saved_profiles.append(entry)
        self._refresh_overlay_labels()

    def _refresh_overlay_labels(self):
        for idx, entry in enumerate(self._saved_profiles, 1):
            label = entry.get('overlay_label_artist')
            if label is not None:
                try:
                    label.set_text(f"Overlay {idx}")
                except Exception:
                    pass

    def clear_saved_profiles(self, notify=True):
        self._clear_saved_profile_artists(notify=notify)

    def highlight_saved_profile(self, index):
        """Update overlay styling to emphasize a selected entry."""
        self._highlighted_overlay = index if index is not None else None
        for idx, entry in enumerate(self._saved_profiles):
            artists = entry.get('artists', [])
            if not artists:
                continue
            line = artists[0]
            base_lw = entry.get('lw', 1.5)
            try:
                if idx == self._highlighted_overlay:
                    line.set_linewidth(base_lw + 1.0)
                    line.set_alpha(1.0)
                else:
                    line.set_linewidth(base_lw)
                    line.set_alpha(0.35)
            except Exception:
                pass
            for label in entry.get('endpoint_labels', []) or []:
                try:
                    label.set_visible(idx == self._highlighted_overlay)
                except Exception:
                    pass
            label_artist = entry.get('overlay_label_artist')
            if label_artist is not None:
                try:
                    label_artist.set_visible(idx == self._highlighted_overlay)
                except Exception:
                    pass
        self.draw_idle()

    def _on_press(self, event):
        if not self.profile_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        shift_pressed = self._shift_pressed(event)
        
        # Right click context menu for profiles
        if event.button == 3:
            # Check overlay first (increased threshold for easier hitting)
            overlay_idx = self._overlay_index_near(x, y, thresh=15.0)
            if overlay_idx is not None:
                self._show_profile_context_menu(event, overlay_idx=overlay_idx)
                return
            # Check active profile
            if self.profile_pts is not None:
                dist_line = self._distance_to_segment_pixels(x, y, self.profile_pts)
                if dist_line <= 15.0:
                    self._show_profile_context_menu(event, active=True)
                    return
            return
        if event.button != 1:
            return
        marker_idx = self._profile_marker_hit(x, y)
        if marker_idx is not None:
            self._profile_marker_drag_idx = marker_idx
            self._dragging = None
            return
        if self.profile_pts is None:
            self._set_profile_pts((x, y, x, y))
            self._ensure_profile_artists()
            self._dragging = 'p1'
            self._update_profile_artists()
            return
        x0, y0, x1, y1 = self.profile_pts
        d0 = self._pt_distance_pixels(x, y, x0, y0)
        d1 = self._pt_distance_pixels(x, y, x1, y1)
        thresh = 18.0  # pixels
        if d0 <= thresh or d0 <= d1:
            if d0 <= thresh:
                self._dragging = 'p0'
                self._line_drag_origin = None
                return
        if d1 <= thresh:
            self._dragging = 'p1'
            self._line_drag_origin = None
            return
        if self.profile_pts is not None:
            dist_line = self._distance_to_segment_pixels(x, y, self.profile_pts)
            if dist_line <= thresh:
                self._dragging = 'line'
                self._line_drag_origin = (x, y, self.profile_pts)
                return
        # Increased threshold to 15.0 to prevent accidental "misses" causing profile loss
        overlay_idx = self._overlay_index_near(x, y, thresh=15.0)
        if overlay_idx is not None:
            if self.profile_pts is not None:
                self._snapshot_active_profile()
            activated = self.activate_saved_profile(overlay_idx)
            if activated:
                if callable(self._profile_highlight_cb):
                    try:
                        self._profile_highlight_cb(None)
                    except Exception:
                        pass
                x0, y0, x1, y1 = self.profile_pts
                return
            else:
                self.highlight_saved_profile(overlay_idx)
                if callable(self._profile_highlight_cb):
                    try:
                        self._profile_highlight_cb(overlay_idx)
                    except Exception:
                        pass
                self._dragging = None
                self._line_drag_origin = None
                return
        # else: start a new line from here
        if self.profile_pts is not None:
            self._snapshot_active_profile()
        self._active_profile_original_color = None
        self._set_profile_pts((x, y, x, y))
        self._dragging = 'p1'
        self._line_drag_origin = None
        
        # Prepare for blitting
        self._set_profile_animated(True)
        self.draw()
        self._profile_background = self.copy_from_bbox(self.main_ax.bbox)
        self._draw_profile_animated()
        self.blit(self.main_ax.bbox)
        self._update_profile_artists()

    def _show_profile_context_menu(self, event, overlay_idx=None, active=False):
        menu = QtWidgets.QMenu(self)
        color_act = menu.addAction("Change Color")
        thicker_act = menu.addAction("Thicker")
        thinner_act = menu.addAction("Thinner")
        
        if overlay_idx is not None:
            menu.addSeparator()
            delete_act = menu.addAction("Delete Profile")
        
        action = menu.exec_(event.guiEvent.globalPos())
        
        if action == color_act:
            self._change_profile_color(overlay_idx, active)
        elif action == thicker_act:
            self._change_profile_width(overlay_idx, active, 0.5)
        elif action == thinner_act:
            self._change_profile_width(overlay_idx, active, -0.5)
        elif overlay_idx is not None and action == delete_act:
            self._remove_saved_profile(overlay_idx)

    def _change_profile_color(self, overlay_idx, active):
        current_color = self._active_profile_color
        if overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles):
            current_color = self._saved_profiles[overlay_idx].get('color', current_color)
        
        col = QtWidgets.QColorDialog.getColor(QtGui.QColor(current_color), self, "Select Profile Color")
        if not col.isValid(): return
        new_color = col.name()
        
        if active:
            self._active_profile_color = new_color
            if self._profile_line:
                self._profile_line.set_color(new_color)
            if self._profile_p0:
                self._profile_p0.set_color(new_color)
            if self._profile_p1:
                self._profile_p1.set_color(new_color)
            for entry in self._profile_echo_artists:
                if entry.get('line'): entry['line'].set_color(new_color)
                if entry.get('p0'): entry['p0'].set_color(new_color)
                if entry.get('p1'): entry['p1'].set_color(new_color)
            self.draw_idle()
            self._emit_profile()
        
        if overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles):
            entry = self._saved_profiles[overlay_idx]
            entry['color'] = new_color
            # Update artists
            for art in entry.get('artists', []):
                try: art.set_color(new_color)
                except: pass
                try: art.set_markeredgecolor('black')
                except: pass
            if entry.get('overlay_label_artist'):
                entry['overlay_label_artist'].set_color(new_color)
            self.draw_idle()
            self._emit_profile()

    def _change_profile_width(self, overlay_idx, active, delta):
        if active:
            self._active_profile_lw = max(0.5, self._active_profile_lw + delta)
            if self._profile_line:
                self._profile_line.set_linewidth(self._active_profile_lw)
                for entry in self._profile_echo_artists:
                    if entry.get('line'): entry['line'].set_linewidth(self._active_profile_lw)
            self.draw_idle()
        
        if overlay_idx is not None and 0 <= overlay_idx < len(self._saved_profiles):
            entry = self._saved_profiles[overlay_idx]
            cur_lw = entry.get('lw', 1.5)
            new_lw = max(0.5, cur_lw + delta)
            entry['lw'] = new_lw
            # Update artists (first artist is usually the line)
            artists = entry.get('artists', [])
            if artists and isinstance(artists[0], Line2D):
                artists[0].set_linewidth(new_lw)
            self.draw_idle()

    def _on_motion(self, event):
        if not self.profile_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        if self._profile_marker_drag_idx is not None:
            if not self._profile_marker_domain or self._profile_marker_positions is None:
                return
            pts = self._profile_marker_pts()
            if pts is None:
                return
            x0, y0, x1, y1 = pts
            vx = x1 - x0
            vy = y1 - y0
            denom = vx * vx + vy * vy
            if denom <= 1e-12:
                return
            t = ((x - x0) * vx + (y - y0) * vy) / denom
            t = max(0.0, min(1.0, t))
            dom_min, dom_max = self._profile_marker_domain
            pos = dom_min + t * (dom_max - dom_min)
            self._profile_marker_positions[self._profile_marker_drag_idx] = pos
            self._update_profile_marker_artists_fast()
            if self._profile_marker_key is not None:
                self._profile_marker_positions_by_key[self._profile_marker_key] = list(self._profile_marker_positions)
            else:
                self._profile_marker_positions_by_key[None] = list(self._profile_marker_positions)
            if callable(self._profile_marker_callback):
                self._profile_marker_callback(list(self._profile_marker_positions), tuple(self._profile_marker_domain))
            self._schedule_profile_update()
            return
        if self._dragging is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        # Hide echo artists during drag for performance
        for entry in self._profile_echo_artists:
            for art in entry.values():
                art.set_visible(False)
        if self._dragging == 'p0':
            self._set_profile_pts((x, y, x1, y1))
        elif self._dragging == 'p1':
            self._set_profile_pts((x0, y0, x, y))
        elif self._dragging == 'line' and self._line_drag_origin is not None and self.profile_pts is not None:
            sx, sy, pts = self._line_drag_origin
            dx = x - sx
            dy = y - sy
            self._set_profile_pts((pts[0] + dx, pts[1] + dy, pts[2] + dx, pts[3] + dy))
        
        # Use blitting for smooth drag
        if self._profile_background:
            self.restore_region(self._profile_background)
            self._update_profile_artists_fast(draw=False)
            self._draw_profile_animated()
            self.blit(self.main_ax.bbox)
        else:
            self._update_profile_artists_fast()
            
        self._schedule_profile_update()

    def _on_release(self, event):
        if not self.profile_enabled:
            return
        self._dragging = None
        self._set_profile_animated(False)
        self._profile_background = None
        # Restore echo artists
        for entry in self._profile_echo_artists:
            for art in entry.values():
                art.set_visible(True)
        self._line_drag_origin = None
        self._profile_marker_drag_idx = None
        self._flush_profile_updates()
        if self._profile_state_deferred:
            self._profile_state_deferred = False
            self._flush_profile_state()

    def _set_profile_animated(self, animated):
        """Set animated state for active profile artists to enable/disable blitting."""
        artists = [self._profile_line, self._profile_p0, self._profile_p1, 
                   self._profile_label, self._profile_ticks, self._profile_info_text]
        artists.extend(self._profile_endpoint_labels)
        for art in artists:
            if art is not None:
                art.set_animated(animated)

    def _draw_profile_animated(self):
        """Draw only the active profile artists (for blitting)."""
        artists = [self._profile_line, self._profile_p0, self._profile_p1, 
                   self._profile_label, self._profile_ticks, self._profile_info_text]
        artists.extend(self._profile_endpoint_labels)
        for art in artists:
            if art is not None and art.get_visible():
                self.main_ax.draw_artist(art)

    def _on_angle_press(self, event):
        if not self.angle_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        if event.button != 1:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        handle = self._angle_handle_at(x, y)
        if handle is None:
            return
        self._angle_dragging = handle
        self._angle_drag_origin = (x, y, self.angle_pts)

    def _on_angle_motion(self, event):
        if not self.angle_enabled or self._angle_dragging is None:
            return
        if event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None or self.angle_pts is None:
            return
        vx, vy, ax, ay, bx, by = self.angle_pts
        if self._angle_dragging == 'vertex':
            dx = x - vx
            dy = y - vy
            self._set_angle_pts(x, y, ax + dx, ay + dy, bx + dx, by + dy)
        elif self._angle_dragging == 'a':
            self._set_angle_pts(vx, vy, x, y, bx, by)
        elif self._angle_dragging == 'b':
            self._set_angle_pts(vx, vy, ax, ay, x, y)
        self._update_angle_artists()
        self._emit_angle()

    def _on_angle_release(self, event):
        if not self.angle_enabled:
            return
        self._angle_dragging = None
        self._angle_drag_origin = None

    def _angle_handle_at(self, x, y):
        if not self.angle_pts:
            return None
        vx, vy, ax, ay, bx, by = self.angle_pts
        handles = {
            'vertex': (vx, vy),
            'a': (ax, ay),
            'b': (bx, by),
        }
        min_handle = None
        min_dist = float('inf')
        for name, (hx, hy) in handles.items():
            dist = self._pt_distance_pixels(x, y, hx, hy)
            if dist < min_dist:
                min_dist = dist
                min_handle = name
        if min_dist <= 12.0:
            return min_handle
        return None
    def _emit_profile(self):
        if not callable(self.profile_callback):
            self._emit_profile_state()
            return
        active = self._build_profile_data(self.profile_pts, color=self._active_profile_color, view=self.views[0] if self.views else None)
        if active:
            ref = active.get('x_nm') if active.get('x_nm') is not None else active.get('x_px')
            if ref is not None:
                try:
                    ref_arr = np.asarray(ref, dtype=float)
                    if ref_arr.size:
                        self._profile_marker_domain = (float(ref_arr.min()), float(ref_arr.max()))
                        if self._profile_marker_key is not None:
                            self._profile_marker_domain_by_key[self._profile_marker_key] = self._profile_marker_domain
                        else:
                            self._profile_marker_domain_by_key[None] = self._profile_marker_domain
                        self._update_profile_marker_artists()
                except Exception:
                    pass
            # Build extra channels if present
            if len(self.views) > 1:
                extras = []
                extra_colors = ['#ff4081', '#00e5ff', '#76ff03', '#d500f9']
                for i, v in enumerate(self.views[1:]):
                    col = extra_colors[i % len(extra_colors)]
                    p = self._build_profile_data(self.profile_pts, color=col, view=v)
                    if p:
                        name = v.get('colorbar_label') or v.get('title') or f"Ch{i+2}"
                        p['name'] = name
                        extras.append(p)
                active['extra_channels'] = extras

        saved_data = []
        for entry in self._saved_profiles:
            data = entry.get('data')
            if data is None:
                data = self._build_profile_data(entry.get('pts'), color=entry.get('color'))
                entry['data'] = data
            if data:
                saved_data.append(data)
        try:
            self.profile_callback(active, saved_data)
        except Exception:
            pass
        self._emit_profile_state()

    def _emit_profile_state(self):
        if self._profile_state_syncing:
            return
        if not callable(self._profile_state_callback):
            return
        if self._dragging is not None or self._profile_marker_drag_idx is not None:
            self._profile_state_deferred = True
            return
        try:
            self._profile_state_callback(self.export_profile_state())
        except Exception:
            pass

    def _flush_profile_state(self):
        if self._profile_state_syncing:
            return
        if not callable(self._profile_state_callback):
            return
        try:
            self._profile_state_callback(self.export_profile_state())
        except Exception:
            pass

    def _on_base_click(self, event):
        if event is None or event.inaxes is None:
            return
        # If clicking on a scale bar, do not trigger base canvas actions (like drag/copy)
        if self.scale_bar_enabled:
            for sb in self._scale_bar_artists:
                if sb.contains(event)[0]:
                    return
        if self.scale_bar_enabled and self._scale_bar_drag_start is not None:
            return
        ax = event.inaxes
        
        if self._check_molecule_hit(event):
            return

        view = self._ax_view_map.get(ax)
        if self.profile_enabled and ax is self.main_ax:
            # avoid starting thumbnail drag/other actions while measuring profiles
            if event.button == 3:
                self._show_context_menu(event, view)
            return
        if self.angle_enabled and ax is self.main_ax:
            if event.button == 3:
                self._show_context_menu(event, view)
            return
        if event.button == 3:
            self._show_context_menu(event, view)
            return
        if event.button != 1:
            return
        if getattr(event, 'dblclick', False):
            if view:
                self._copy_view_to_clipboard(view)
            return
        if view and getattr(event, 'guiEvent', None) is not None:
            pos = event.guiEvent.globalPos()
            self._drag_candidate = {'view': view, 'start': QtCore.QPoint(pos), 'image': None}

    def _load_molecule_dialog(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Molecule", "", "Molecule Files (*.xyz *.pdb *.mol);;All Files (*)"
        )
        if path:
            self.add_molecule(path)

    def add_molecule(self, path):
        try:
            mol = Molecule(path)
            # Center in current view if possible
            if self.main_ax:
                xlim = self.main_ax.get_xlim()
                ylim = self.main_ax.get_ylim()
                mol.offset = np.array([(xlim[0]+xlim[1])/2, (ylim[0]+ylim[1])/2, 0.0])
            self.molecules.append(mol)
            self._redraw()
        except Exception as e:
            print(f"Failed to load molecule: {e}")

    def _clear_molecules(self):
        self.molecules = []
        self._redraw()

    def _check_molecule_hit(self, event):
        if not self.molecules or event.inaxes is None:
            return False
        
        # Simple hit test: check distance to any atom in any molecule
        # Iterate in reverse to pick top-most
        for idx, mol in reversed(list(enumerate(self.molecules))):
            coords = mol.get_transformed_coordinates()
            if len(coords) == 0: continue
            
            # Map event data coords to display coords is hard without transform
            # We'll just check data distance. 
            # Assuming atoms are roughly 0.1-0.3 nm radius visually
            dx = coords[:, 0] - event.xdata
            dy = coords[:, 1] - event.ydata
            dist_sq = dx*dx + dy*dy
            min_dist = np.min(dist_sq)
            
            # Threshold: 0.5 nm radius click tolerance
            if min_dist < 0.25: 
                if event.button == 1:
                    self._molecule_drag_idx = idx
                    self._molecule_drag_start = (event.xdata, event.ydata)
                    self._molecule_drag_start_px = (event.x, event.y)
                    self._molecule_drag_mol_start = mol.offset.copy()
                    self._molecule_drag_mol_angles = mol.angles.copy()
                    
                    key = str(event.key).lower() if event.key else ''
                    if 'control' in key and 'shift' in key:
                        self._molecule_drag_mode = 'rotate_3d'
                    elif 'shift' in key:
                        self._molecule_drag_mode = 'rotate_z'
                    else:
                        self._molecule_drag_mode = 'translate'
                        
                    return True
                elif event.button == 3:
                    self._show_molecule_menu(event, mol)
                    return True
        return False

    def _on_molecule_motion(self, event):
        if self._molecule_drag_idx is not None:
            if event.xdata is None or event.ydata is None:
                return
            mol = self.molecules[self._molecule_drag_idx]
            
            if self._molecule_drag_mode == 'translate':
                dx = event.xdata - self._molecule_drag_start[0]
                dy = event.ydata - self._molecule_drag_start[1]
                mol.offset = self._molecule_drag_mol_start + np.array([dx, dy, 0.0])
            
            elif self._molecule_drag_mode == 'rotate_z':
                center = self._molecule_drag_mol_start
                v_start = np.array([self._molecule_drag_start[0] - center[0], self._molecule_drag_start[1] - center[1]])
                v_curr = np.array([event.xdata - center[0], event.ydata - center[1]])
                if np.linalg.norm(v_start) > 0.01 and np.linalg.norm(v_curr) > 0.01:
                    angle_start = np.arctan2(v_start[1], v_start[0])
                    angle_curr = np.arctan2(v_curr[1], v_curr[0])
                    delta_deg = np.degrees(angle_curr - angle_start)
                    new_angles = self._molecule_drag_mol_angles.copy()
                    new_angles[2] += delta_deg
                    mol.angles = new_angles
            
            elif self._molecule_drag_mode == 'rotate_3d':
                if event.x is None or event.y is None: return
                dx_px = event.x - self._molecule_drag_start_px[0]
                dy_px = event.y - self._molecule_drag_start_px[1]
                sensitivity = 0.5 # degrees per pixel
                new_angles = self._molecule_drag_mol_angles.copy()
                new_angles[0] += dy_px * sensitivity
                new_angles[1] += dx_px * sensitivity
                mol.angles = new_angles

            # Update rotation guide (visual only, no full redraw needed)
            if self._molecule_drag_mode in ('rotate_z', 'rotate_3d'):
                if self._molecule_rotation_guide is None and self.main_ax:
                    self._molecule_rotation_guide = patches.Circle(
                        (mol.offset[0], mol.offset[1]), 
                        radius=2.0, # Fixed visual radius or dynamic based on molecule size
                        fill=False, edgecolor='yellow', linestyle='--', linewidth=1.5, alpha=0.6, zorder=40
                    )
                    self.main_ax.add_patch(self._molecule_rotation_guide)
                elif self._molecule_rotation_guide:
                    self._molecule_rotation_guide.center = (mol.offset[0], mol.offset[1])

            self._update_molecule_artists()

    def _on_molecule_release(self, event):
        if self._molecule_drag_idx is not None:
            self._molecule_drag_idx = None
            self._molecule_drag_start = None
            self._molecule_drag_start_px = None
            self._molecule_drag_mode = None
            self._molecule_drag_mol_angles = None
            
            if self._molecule_rotation_guide:
                try: self._molecule_rotation_guide.remove()
                except: pass
                self._molecule_rotation_guide = None
                self._redraw()

    def _show_molecule_menu(self, event, mol):
        menu = QtWidgets.QMenu(self)
        props_act = menu.addAction("Properties (Rotate/Scale)...")
        dup_act = menu.addAction("Duplicate")
        del_act = menu.addAction("Delete")
        
        action = menu.exec_(event.guiEvent.globalPos())
        if action == props_act:
            dlg = MoleculePropertiesDialog(mol, self, callback=self._redraw)
            dlg.show()
        elif action == dup_act:
            new_mol = mol.copy()
            new_mol.offset += np.array([1.0, 1.0, 0.0]) # Slight offset
            self.molecules.append(new_mol)
            self._redraw()
        elif action == del_act:
            if mol in self.molecules:
                self.molecules.remove(mol)
                self._redraw()

    def _copy_view_to_clipboard(self, view):
        try:
            qimg = self._view_to_qimage(view)
            QtWidgets.QApplication.clipboard().setImage(qimg)
            if callable(self._copy_feedback_handler):
                self._copy_feedback_handler(view)
        except Exception:
            pass

    def _view_to_qimage(self, view):
        arr = np.asarray(view.get('arr'))
        cmap = view.get('cmap', 'viridis')
        return array_to_qimage(arr, cmap_name=cmap)

    def _show_context_menu(self, event, view):
        if view is None or getattr(event, 'guiEvent', None) is None:
            return
        menu = QtWidgets.QMenu(self)
        copy_act = menu.addAction("Copy image")
        copy_svg_act = menu.addAction("Copy view as SVG (vector)")
        copy_disp_png = menu.addAction("Copy displayed (PNG)")
        copy_disp_svg = menu.addAction("Copy displayed (SVG)")
        save_act = menu.addAction("Save image as...")
        save_svg_act = menu.addAction("Save view as SVG...")
        save_pdf_act = menu.addAction("Save view as PDF...")
        
        menu.addSeparator()
        reset_zoom_act = menu.addAction("Reset Zoom")
        
        cbar_text = "Horizontal Colorbar" if self._colorbar_orientation == 'vertical' else "Vertical Colorbar"
        toggle_cbar_act = menu.addAction(cbar_text)

        menu.addSeparator()
        show_ticks_act = menu.addAction("Show Ticks")
        show_ticks_act.setCheckable(True)
        show_ticks_act.setChecked(self._show_ticks)
        show_cbar_act = menu.addAction("Show Colorbar")
        show_cbar_act.setCheckable(True)
        show_cbar_act.setChecked(self._show_colorbar)
        
        menu.addSeparator()
        load_mol_act = menu.addAction("Load Molecule (XYZ/PDB)...")
        clear_mols_act = menu.addAction("Clear Molecules")

        chosen = menu.exec_(event.guiEvent.globalPos())
        if chosen == copy_act:
            self._copy_view_to_clipboard(view)
        elif chosen == copy_svg_act:
            self._copy_view_as_svg(view)
        elif chosen == copy_disp_png:
            self._copy_displayed("png")
        elif chosen == copy_disp_svg:
            self._copy_displayed("svg")
        elif chosen == save_act:
            self._save_view_to_file(view)
        elif chosen == save_svg_act:
            self._save_view_vector(view, "svg")
        elif chosen == save_pdf_act:
            self._save_view_vector(view, "pdf")
        elif chosen == reset_zoom_act:
            self._reset_view_zoom()
        elif chosen == toggle_cbar_act:
            self._toggle_colorbar_orientation()
        elif chosen == show_ticks_act:
            self._toggle_ticks()
        elif chosen == show_cbar_act:
            self._toggle_colorbar()
        elif chosen == load_mol_act:
            self._load_molecule_dialog()
        elif chosen == clear_mols_act:
            self._clear_molecules()

    def _save_view_to_file(self, view):
        try:
            title = view.get('title') or 'view'
            default = f"{title}.png"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save view", default, "PNG Files (*.png)")
            if not path:
                return
            qimg = self._view_to_qimage(view)
            qimg.save(path, "PNG")
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Save view", "Unable to save image.")

    def _copy_displayed(self, fmt='png'):
        """Copy the current figure exactly as displayed (including overlays)."""
        buf = io.BytesIO()
        if fmt == 'svg':
            with matplotlib.rc_context({'svg.fonttype': 'none'}):
                self.fig.savefig(buf, format='svg', bbox_inches='tight')
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", buf.getvalue())
            QtWidgets.QApplication.clipboard().setMimeData(mime)
        else:
            self.fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
            qimg = QtGui.QImage.fromData(buf.getvalue())
            QtWidgets.QApplication.clipboard().setImage(qimg)

    def _copy_view_as_svg(self, view):
        try:
            fig = self._render_view_figure(view)
            buf = io.BytesIO()
            with matplotlib.rc_context({'svg.fonttype': 'none'}):
                fig.savefig(buf, format="svg", bbox_inches="tight", pad_inches=0.02)
            svg_bytes = buf.getvalue()
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", svg_bytes)
            QtWidgets.QApplication.clipboard().setMimeData(mime)
            if callable(self._copy_feedback_handler):
                self._copy_feedback_handler(view)
        except Exception:
            pass

    def _save_view_vector(self, view, fmt):
        fmt = (fmt or "").strip().lower()
        if fmt not in ("svg", "pdf"):
            return
        try:
            title = view.get('title') or 'view'
            default = f"{title}.{fmt}"
            label = "SVG Files (*.svg)" if fmt == "svg" else "PDF Files (*.pdf)"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save view", default, label)
            if not path:
                return
            if not path.lower().endswith(f".{fmt}"):
                path = f"{path}.{fmt}"
            fig = self._render_view_figure(view)
            if fmt == 'svg':
                with matplotlib.rc_context({'svg.fonttype': 'none'}):
                    fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            else:
                fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            try:
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))
            except Exception:
                pass
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Save view", "Unable to save vector image.")

    def _reset_view_zoom(self):
        self._view_font_scale = 1.0
        self._apply_view_font_scale()

    def _toggle_colorbar_orientation(self):
        self._colorbar_orientation = 'horizontal' if self._colorbar_orientation == 'vertical' else 'vertical'
        self._redraw()

    def _toggle_ticks(self):
        self._show_ticks = not self._show_ticks
        self._redraw()

    def _toggle_colorbar(self):
        self._show_colorbar = not self._show_colorbar
        self._redraw()

    def _render_view_figure(self, view):
        fig = Figure(figsize=(6, 6))
        ax = fig.add_subplot(1, 1, 1)
        arr = np.asarray(view.get('arr'))
        flip = bool(view.get('relative_axes'))
        if flip:
            arr_plot = np.flipud(arr)
        else:
            arr_plot = arr
        extent = view.get('extent', None)
        cmap = view.get('cmap', 'viridis')
        origin = 'lower' if flip else 'upper'
        if extent is None:
            im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
        else:
            im = ax.imshow(arr_plot, extent=extent, origin=origin, interpolation='nearest', aspect='equal', cmap=cmap)
        ax.set_autoscale_on(False)
        cbar_label = view.get('colorbar_label') or view.get('unit', '')
        cbar = None
        if cbar_label and self._show_colorbar:
            try:
                divider = make_axes_locatable(ax)
                if self._colorbar_orientation == 'horizontal':
                    cax = divider.append_axes("bottom", size="5%", pad=0.08)
                    cbar = fig.colorbar(im, cax=cax, orientation='horizontal')
                    cbar.set_label(cbar_label)
                    cbar.ax.xaxis.set_label_coords(0.5, 0.5)
                    cbar.ax.xaxis.label.set_horizontalalignment('center')
                    cbar.ax.xaxis.label.set_verticalalignment('center')
                else:
                    cax = divider.append_axes("right", size="4%", pad=0.02)
                    cbar = fig.colorbar(im, cax=cax, orientation='vertical')
                    cbar.set_label(cbar_label)
                    cbar.ax.yaxis.set_label_coords(0.5, 0.5)
                    cbar.ax.yaxis.label.set_horizontalalignment('center')
                    cbar.ax.yaxis.label.set_verticalalignment('center')
            except Exception:
                cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02, orientation=self._colorbar_orientation)
                cbar.set_label(cbar_label)
            if not self._show_ticks:
                cbar.set_ticks([])
        title = view.get('title', '')
        if title:
            ax.set_title(title, fontsize=9)
        ax.tick_params(labelsize=8)

        if self.scale_bar_enabled:
            extent = view.get('extent')
            if extent is None:
                h, w = np.shape(view['arr'])
                width = w
                unit = 'px'
            else:
                width = abs(extent[1] - extent[0])
                unit = view.get('axis_unit') or 'nm'
            
            size, label = self._calculate_best_scale_bar(width, unit)
            font_scale = getattr(self, '_view_font_scale', 1.0)
            
            dark = bool(self._detail_dark)
            default_color = '#f5f5f5' if dark else '#111111'
            sb_settings = getattr(self, '_scale_bar_settings', {})
            sb_text_col = sb_settings.get('text_color') or default_color
            sb_bar_col = sb_settings.get('bar_color') or default_color
            font_family = sb_settings.get('font_family', 'sans-serif')

            sb = AnchoredSizeBar(ax.transData, size, label, loc='center',
                                 pad=0.4, borderpad=0, sep=3, frameon=False,
                                 size_vertical=width*0.004*font_scale, color=sb_bar_col,
                                 label_top=True,
                                 bbox_to_anchor=self._scale_bar_pos, bbox_transform=ax.transAxes)
            sb.size_bar.get_children()[0].set_linewidth(0)
            text = sb.txt_label.get_children()[0]
            text.set_color(sb_text_col)
            text.set_fontfamily(font_family)
            text.set_fontsize(10 * font_scale)
            text.set_fontweight('bold')
            ax.add_artist(sb)

        if not self._show_ticks:
            ax.set_xticks([])
            ax.set_yticks([])

        self._draw_molecules(ax)

        self._style_export_figure(fig, ax, cbar)
        try:
            fig.tight_layout()
        except Exception:
            pass
        return fig

    def _style_export_figure(self, fig, ax, cbar):
        dark = bool(self._detail_dark)
        fig_face = '#111217' if dark else '#ffffff'
        ax_face = '#14161c' if dark else '#ffffff'
        text_color = '#f5f5f5' if dark else '#111111'
        grid_color = '#4f5a64' if dark else '#9a9a9a'
        try:
            fig.set_facecolor(fig_face)
        except Exception:
            pass
        try:
            ax.set_facecolor(ax_face)
            ax.tick_params(colors=text_color, labelcolor=text_color)
            ax.xaxis.label.set_color(text_color)
            ax.yaxis.label.set_color(text_color)
            for spine in ax.spines.values():
                spine.set_color(text_color)
            if self._detail_grid:
                ax.grid(True, color=grid_color, alpha=0.3, linewidth=0.6)
            else:
                ax.grid(False)
        except Exception:
            pass
        if cbar is not None:
            try:
                cbar.ax.tick_params(colors=text_color, labelcolor=text_color)
                cbar.ax.yaxis.label.set_color(text_color)
                cbar.ax.xaxis.label.set_color(text_color)
                cbar.outline.set_edgecolor(text_color)
            except Exception:
                pass
        scale = max(0.6, min(2.5, getattr(self, '_view_font_scale', 1.0)))
        tick_size = 8 * scale
        label_size = 10 * scale
        title_size = 9 * scale
        try:
            ax.tick_params(labelsize=tick_size)
            ax.xaxis.label.set_fontsize(label_size)
            ax.yaxis.label.set_fontsize(label_size)
            ax.title.set_fontsize(title_size)
        except Exception:
            pass
        if cbar is not None:    
            try:
                cbar.ax.tick_params(labelsize=tick_size)
                cbar.ax.yaxis.label.set_fontsize(label_size)
                cbar.ax.xaxis.label.set_fontsize(label_size)
            except Exception:
                pass

    def _start_drag(self, view, qimg=None):
        try:
            if qimg is None:
                qimg = self._view_to_qimage(view)
            pix = QtGui.QPixmap.fromImage(qimg)
            drag = QtGui.QDrag(self)
            mime = QtCore.QMimeData()
            mime.setImageData(qimg)
            try:
                meta = view.get('meta') or {}
                payload = {
                    'file_path': meta.get('file_path'),
                    'channel_index': meta.get('channel_index'),
                    'cmap': view.get('cmap'),
                }
                if payload.get('file_path') is not None and payload.get('channel_index') is not None:
                    mime.setData('application/x-sxm-view', json.dumps(payload).encode('utf-8'))
            except Exception:
                pass
            drag.setMimeData(mime)
            drag.setPixmap(pix.scaled(128, 128, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
            drag.exec_(QtCore.Qt.CopyAction)
        except Exception:
            pass

    def mouseMoveEvent(self, event):
        if self._drag_candidate:
            start = self._drag_candidate.get('start')
            if start is not None:
                if (event.globalPos() - start).manhattanLength() >= 10:
                    view = self._drag_candidate.get('view')
                    qimg = self._drag_candidate.get('image')
                    if qimg is None and view is not None:
                        qimg = self._view_to_qimage(view)
                        self._drag_candidate['image'] = qimg
                    if view is not None and qimg is not None:
                        self._start_drag(view, qimg)
                    self._drag_candidate = None
                    super().mouseMoveEvent(event)
                    return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_candidate = None
        super().mouseReleaseEvent(event)

    def _on_motion_value(self, event):
        if self._value_callback is None:
            return
        # Performance: skip value inspection while dragging profiles or molecules
        if getattr(self, '_dragging', None) is not None or getattr(self, '_molecule_drag_idx', None) is not None:
            return
            
        if event.inaxes is None or event.inaxes not in self._ax_view_map:
            self._value_callback(None, None, None, None)
            return
        view = self._ax_view_map.get(event.inaxes)
        if view is None:
            self._value_callback(None, None, None, None)
            return
        val = sample_array_value(view.get('arr'), event.xdata, event.ydata, view.get('extent'))
        if val is None:
            self._value_callback(None, event.xdata, event.ydata, view)
        else:
            self._value_callback(val, event.xdata, event.ydata, view)

class SafeFigureCanvas(FigureCanvas):
    def draw(self):
        try:
            super().draw()
        except np.linalg.LinAlgError:
            # Ignore transient singular transforms during layout updates.
            return
