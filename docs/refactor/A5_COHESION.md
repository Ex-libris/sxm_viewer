# A5 - Attribute cohesion (hidden class boundaries)

Clusters a large class's attributes by which methods use them. Attributes that always travel together are one piece of state - a candidate extracted class. Isolation scores rank which groups can be pulled out with the least ripple. Read-only evidence; no refactoring is implied by inclusion here.

## SXMGridViewer

`sxm_viewer/gui/main_window.py`

- methods analysed: **537**
- distinct `self.` attributes: **814**
- cohesive groups (>= 3 attrs): **30**

### Candidate extractions, most isolated first

`Isolation` = share of methods touching this group that touch *nothing else*. High = safe to extract.

| Isolation | Attrs | Pure/Touching methods | Attributes |
|---|---|---|---|
| 75% | 3 | 3/4 | `_activity_log_pending`, `_activity_log_flush_timer`, `activity_log_box` |
| 50% | 4 | 1/2 | `_session_activity_title`, `_session_activity_strip`, `_session_activity_detail`, `_session_activity_progress` |
| 50% | 3 | 4/8 | `_suspend_window_history`, `_capture_window_state_payload`, `_push_closed_window_history` |
| 50% | 3 | 2/4 | `spec_folder_le`, `spec_folder_path`, `_set_spec_folder` |
| 40% | 3 | 2/5 | `MODE_BROWSE`, `MODE_SPECTRO`, `MODE_MEASURE` |
| 40% | 3 | 2/5 | `_apply_ui_theme`, `_sync_forced_cmap`, `ui_theme` |
| 33% | 4 | 1/3 | `_pending_preview_request`, `_flush_preview_request`, `_preview_request_timer`, `_preview_render_in_progress` |
| 33% | 4 | 1/3 | `on_recall_popouts`, `on_arrange_popouts`, `on_close_popouts`, `on_minimize_popouts` |
| 33% | 3 | 1/3 | `_thumbnail_render_state_pending_paths`, `_thumbnail_render_state_timer`, `_flush_thumbnail_render_state_refresh` |
| 25% | 19 | 1/4 | `virtual_copy_order`, `_deferred_popup_serial`, `_last_base_array`, `spectro_sites_by_image`, `_hide_session_activity`, `_spectro_hist_cache` +13 |
| 25% | 3 | 1/4 | `open_spectro_browser`, `_set_spectro_browser_filters`, `_spec_matches_image_key` |
| 25% | 3 | 1/4 | `auto_detect_tags`, `_workspace_loading`, `_auto_detect_tags_for_folder` |
| 20% | 5 | 1/5 | `spectros_by_image`, `files_with_spectra`, `_spec_extent_cache`, `files_with_matrix`, `matrix_datasets` |
| 14% | 6 | 1/7 | `main_splitter`, `_update_preview_detach_button`, `_preview_panel`, `_layout_sizes`, `preview_detached`, `_preview_dialog` |
| 14% | 4 | 1/7 | `_get_adjust_spec`, `_adjustment_undo_stack`, `_set_adjust_spec`, `_refresh_adjusted_channel` |
| 12% | 8 | 1/8 | `_thumb_inflight`, `_request_visible_thumbs`, `_thumb_loaded`, `_thumb_generation`, `_thumb_labels`, `_thumb_crop_cache` +2 |
| 0% | 7 | 0/6 | `scale_bar_cb`, `show_acquisition_overlay`, `show_molecules`, `_canvas_display_syncing`, `show_molecule_gizmo`, `canvas_display_options` +1 |
| 0% | 6 | 0/3 | `_pending_compact_histogram_clim`, `_compact_histogram_gesture_active`, `_pending_compact_histogram_final`, `_flush_compact_histogram_clim`, `_compact_histogram_apply_timer`, `_suppress_compact_histogram_refresh` |
| 0% | 5 | 0/4 | `_highlighted_spec`, `_highlight_pulse_strength`, `_highlight_phase`, `_highlight_timer`, `_on_highlight_tick` |
| 0% | 4 | 0/4 | `on_show_molecules_toggled`, `_on_recent_molecules_updated`, `on_load_molecule`, `_on_molecule_palette_changed` |
| 0% | 4 | 0/4 | `starred`, `on_add_selected_thumbnails_to_collection`, `report_controller`, `on_adjust_image` |
| 0% | 4 | 0/5 | `thumb_cmap_combo`, `thumb_cmap`, `frame_real_view`, `_thumbnail_cmap_override` |
| 0% | 4 | 0/2 | `_plot_font_italic`, `_plot_font_family`, `_plot_font_underline`, `_plot_font_bold` |
| 0% | 3 | 0/6 | `_filtered_channel_cache`, `_frame_real_pixmap_cache`, `_filtered_cache_lock` |
| 0% | 3 | 0/4 | `spectro_thumb_channel_by_path`, `spectro_miniature_default_channel`, `spectro_compare_controller` |

### Group detail

<details><summary>3 attributes, 75% isolated (3/4 methods)</summary>

- `self._activity_log_flush_timer` - 2 methods
- `self._activity_log_pending` - 4 methods
- `self.activity_log_box` - 2 methods

</details>

<details><summary>4 attributes, 50% isolated (1/2 methods)</summary>

- `self._session_activity_detail` - 2 methods
- `self._session_activity_progress` - 2 methods
- `self._session_activity_strip` - 2 methods
- `self._session_activity_title` - 2 methods

</details>

<details><summary>3 attributes, 50% isolated (4/8 methods)</summary>

- `self._capture_window_state_payload` - 5 methods
- `self._push_closed_window_history` - 5 methods
- `self._suspend_window_history` - 8 methods

</details>

<details><summary>3 attributes, 50% isolated (2/4 methods)</summary>

- `self._set_spec_folder` - 2 methods
- `self.spec_folder_le` - 3 methods
- `self.spec_folder_path` - 3 methods

</details>

<details><summary>3 attributes, 40% isolated (2/5 methods)</summary>

- `self.MODE_BROWSE` - 5 methods
- `self.MODE_MEASURE` - 4 methods
- `self.MODE_SPECTRO` - 4 methods

</details>

<details><summary>3 attributes, 40% isolated (2/5 methods)</summary>

- `self._apply_ui_theme` - 3 methods
- `self._sync_forced_cmap` - 3 methods
- `self.ui_theme` - 3 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self._flush_preview_request` - 3 methods
- `self._pending_preview_request` - 3 methods
- `self._preview_render_in_progress` - 2 methods
- `self._preview_request_timer` - 3 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self.on_arrange_popouts` - 2 methods
- `self.on_close_popouts` - 2 methods
- `self.on_minimize_popouts` - 2 methods
- `self.on_recall_popouts` - 3 methods

</details>

<details><summary>3 attributes, 33% isolated (1/3 methods)</summary>

- `self._flush_thumbnail_render_state_refresh` - 2 methods
- `self._thumbnail_render_state_pending_paths` - 3 methods
- `self._thumbnail_render_state_timer` - 2 methods

</details>

<details><summary>19 attributes, 25% isolated (1/4 methods)</summary>

- `self._collection_item_snapshots` - 2 methods
- `self._deferred_popup_entries` - 2 methods
- `self._deferred_popup_serial` - 2 methods
- `self._hide_session_activity` - 2 methods
- `self._last_base_array` - 2 methods
- `self._last_base_extent` - 2 methods
- `self._last_base_unit` - 2 methods
- `self._refresh_deferred_popup_ui` - 2 methods
- `self._spectro_hist_cache` - 2 methods
- `self._update_toolbar_actions` - 2 methods
- `self.angle_value_label` - 2 methods
- `self.current_spectro_thumb_files` - 2 methods
- `self.preview_value_label` - 2 methods
- `self.selected_spectro_thumb_file` - 2 methods
- `self.spectro_group_index` - 2 methods
- `self.spectro_groups_by_image` - 2 methods
- `self.spectro_site_index` - 2 methods
- `self.spectro_sites_by_image` - 2 methods
- `self.virtual_copy_order` - 4 methods

</details>

<details><summary>3 attributes, 25% isolated (1/4 methods)</summary>

- `self._set_spectro_browser_filters` - 2 methods
- `self._spec_matches_image_key` - 2 methods
- `self.open_spectro_browser` - 4 methods

</details>

<details><summary>3 attributes, 25% isolated (1/4 methods)</summary>

- `self._auto_detect_tags_for_folder` - 3 methods
- `self._workspace_loading` - 3 methods
- `self.auto_detect_tags` - 4 methods

</details>

### Spine attributes (touched by the most methods)

> These resist extraction and define what the class is *actually* about; everything else is a passenger.

| Attribute | Methods touching |
|---|---|
| `self.config` | 72 |
| `self.last_preview` | 44 |
| `self.channel_dropdown` | 38 |
| `self.show_file_channel` | 38 |
| `self.filter_controller` | 31 |
| `self.populate_thumbnails_for_channel` | 29 |
| `self._spectros_loaded` | 22 |
| `self.collection_controller` | 19 |
| `self.preview_canvas` | 18 |
| `self.headers` | 18 |
| `self._schedule_marker_refresh` | 16 |
| `self.ensure_spectros_loaded` | 14 |
| `self.session_controller` | 11 |
| `self._is_processed_key` | 11 |
| `self._processed_views` | 10 |

## MultiPreviewCanvas

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- methods analysed: **439**
- distinct `self.` attributes: **603**
- cohesive groups (>= 3 attrs): **26**

### Candidate extractions, most isolated first

`Isolation` = share of methods touching this group that touch *nothing else*. High = safe to extract.

| Isolation | Attrs | Pure/Touching methods | Attributes |
|---|---|---|---|
| 88% | 20 | 14/16 | `_collection_help_callback`, `_apply_popup_style_callback`, `_close_windows_callback`, `_compare_menu_callback`, `_histogram_dialog_callback`, `_display_relative_zero_menu_state_callback` +14 |
| 43% | 3 | 3/7 | `_recent_svg_molecule_paths`, `_last_svg_molecule_dir_cb`, `_last_svg_molecule_dir` |
| 33% | 9 | 2/6 | `_colorbars`, `_hover_view_ax`, `_active_view_ax`, `_suppress_internal_draw_requests`, `_compute_theme_sig`, `_draw_spectra` +3 |
| 33% | 4 | 1/3 | `_resize_draft_timer`, `_resize_settle_timer`, `_last_resize_size`, `_resize_reflow_threshold_px` |
| 29% | 8 | 2/7 | `_outline_start`, `_outline_ax`, `_crop_square`, `_crop_rect`, `_crop_ax`, `_crop_start` +2 |
| 25% | 4 | 1/4 | `_histogram_auto_callback`, `_reset_view_zoom`, `_load_molecule_dialog`, `_copy_displayed` |
| 14% | 3 | 1/7 | `_relative_axes_override`, `_profile_label_mode`, `_show_molecule_gizmo` |
| 4% | 4 | 1/24 | `_fixed_crop_template`, `_fixed_crop_template_bounds`, `_fixed_crop_template_view_key`, `_fixed_crop_template_pixel_bounds` |
| 0% | 14 | 0/7 | `_svg_molecule_drag`, `_molecule_drag_idx`, `_pan_active`, `_pan_start`, `_molecule_gizmo_drag`, `_molecule_drag_start_px` +8 |
| 0% | 13 | 0/4 | `set_show_shortcut_hint`, `set_show_profile_overlays`, `enable_scale_bar`, `_on_apply_fixed_crop_shortcut`, `set_show_angle_overlays`, `set_show_acquisition_overlay` +7 |
| 0% | 6 | 0/15 | `_show_title`, `_show_ticks`, `_show_colorbar`, `_colorbar_orientation`, `_fit_to_canvas`, `_frame_fill_mode` |
| 0% | 6 | 0/12 | `_profile_p1`, `_profile_line`, `_profile_echo_artists`, `_profile_p0`, `_profile_endpoint_labels`, `_profile_label` |
| 0% | 6 | 0/4 | `_draw_acquisition_overlay`, `_draw_image_size_overlay`, `_draw_outlines`, `_draw_filter_summary_overlay`, `_compose_view_title`, `_draw_molecules` |
| 0% | 5 | 0/17 | `_active_profile_color`, `_active_profile_lw`, `_active_profile_marker_style`, `_active_profile_line_style`, `_active_profile_marker_size` |
| 0% | 4 | 0/15 | `_show_shortcut_hint`, `_show_acquisition_overlay`, `_show_profile_overlays`, `_show_angle_overlays` |
| 0% | 4 | 0/7 | `_dragging`, `_profile_marker_drag_idx`, `_saved_profile_drag`, `_line_drag_origin` |
| 0% | 4 | 0/8 | `_undo_suspend_depth`, `_ensure_angle_frames`, `_undo_history`, `_undo_restore_in_progress` |
| 0% | 4 | 0/6 | `_molecule_gizmo_axes`, `_molecule_gizmo_timer`, `_molecule_gizmo_until`, `_molecule_gizmo_artists` |
| 0% | 4 | 0/7 | `_refresh_overlay_labels`, `_create_profile_id_label`, `_create_endpoint_labels`, `_create_ticks_and_label` |
| 0% | 4 | 0/4 | `_outline_default_ls`, `_outline_default_color`, `_outline_default_lw`, `_outline_threshold` |
| 0% | 4 | 0/6 | `_apply_view_theme`, `_update_highlight_artists`, `_apply_view_font_scale`, `_add_scale_bar` |
| 0% | 3 | 0/12 | `_profile_marker_key`, `_profile_marker_positions_by_key`, `_profile_marker_domain_by_key` |
| 0% | 3 | 0/6 | `_show_molecule_shadow`, `molecule_palette`, `_show_hydrogens` |
| 0% | 3 | 0/4 | `_svg_molecule_drag_background`, `_svg_molecule_entry_for`, `_svg_molecule_blit_artists` |
| 0% | 3 | 0/4 | `_profile_color_cycle`, `_profile_palette_name`, `_profile_palette_colors` |

### Group detail

<details><summary>20 attributes, 88% isolated (14/16 methods)</summary>

- `self._apply_popup_style_callback` - 3 methods
- `self._apply_popup_style_label` - 3 methods
- `self._apply_popup_style_tooltip` - 3 methods
- `self._arrange_windows_callback` - 3 methods
- `self._close_windows_callback` - 3 methods
- `self._collection_help_callback` - 3 methods
- `self._collection_menu_callback` - 3 methods
- `self._compare_menu_callback` - 3 methods
- `self._compare_menu_state_callback` - 3 methods
- `self._display_relative_zero_menu_callback` - 3 methods
- `self._display_relative_zero_menu_state_callback` - 3 methods
- `self._display_relative_zero_menu_tooltip` - 3 methods
- `self._filter_menu_callback` - 3 methods
- `self._histogram_dialog_callback` - 3 methods
- `self._histogram_reset_callback` - 3 methods
- `self._minimize_windows_callback` - 3 methods
- `self._restore_windows_callback` - 3 methods
- `self._spectra_compare_all_callback` - 3 methods
- `self._stp_export_callback` - 3 methods
- `self._virtual_copy_callback` - 3 methods

</details>

<details><summary>3 attributes, 43% isolated (3/7 methods)</summary>

- `self._last_svg_molecule_dir` - 4 methods
- `self._last_svg_molecule_dir_cb` - 4 methods
- `self._recent_svg_molecule_paths` - 5 methods

</details>

<details><summary>9 attributes, 33% isolated (2/6 methods)</summary>

- `self._active_view_ax` - 4 methods
- `self._async_redraw_once` - 2 methods
- `self._colorbars` - 4 methods
- `self._compute_theme_sig` - 2 methods
- `self._draw_fixed_crop_history` - 2 methods
- `self._draw_shortcut_hint` - 2 methods
- `self._draw_spectra` - 2 methods
- `self._hover_view_ax` - 4 methods
- `self._suppress_internal_draw_requests` - 2 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self._last_resize_size` - 2 methods
- `self._resize_draft_timer` - 3 methods
- `self._resize_reflow_threshold_px` - 2 methods
- `self._resize_settle_timer` - 3 methods

</details>

<details><summary>8 attributes, 29% isolated (2/7 methods)</summary>

- `self._crop_ax` - 5 methods
- `self._crop_last_ts` - 3 methods
- `self._crop_rect` - 5 methods
- `self._crop_square` - 5 methods
- `self._crop_start` - 5 methods
- `self._outline_ax` - 6 methods
- `self._outline_rect` - 4 methods
- `self._outline_start` - 6 methods

</details>

<details><summary>4 attributes, 25% isolated (1/4 methods)</summary>

- `self._copy_displayed` - 2 methods
- `self._histogram_auto_callback` - 4 methods
- `self._load_molecule_dialog` - 2 methods
- `self._reset_view_zoom` - 2 methods

</details>

<details><summary>3 attributes, 14% isolated (1/7 methods)</summary>

- `self._profile_label_mode` - 4 methods
- `self._relative_axes_override` - 5 methods
- `self._show_molecule_gizmo` - 4 methods

</details>

<details><summary>4 attributes, 4% isolated (1/24 methods)</summary>

- `self._fixed_crop_template` - 22 methods
- `self._fixed_crop_template_bounds` - 17 methods
- `self._fixed_crop_template_pixel_bounds` - 11 methods
- `self._fixed_crop_template_view_key` - 13 methods

</details>

<details><summary>14 attributes, 0% isolated (0/7 methods)</summary>

- `self._molecule_drag_idx` - 4 methods
- `self._molecule_drag_mode` - 4 methods
- `self._molecule_drag_mol_angles` - 4 methods
- `self._molecule_drag_start` - 4 methods
- `self._molecule_drag_start_px` - 4 methods
- `self._molecule_gizmo_drag` - 4 methods
- `self._molecule_rotation_guide` - 3 methods
- `self._pan_active` - 4 methods
- `self._pan_ax` - 4 methods
- `self._pan_last_ts` - 3 methods
- `self._pan_start` - 4 methods
- `self._pan_start_lim` - 4 methods
- `self._pan_throttle_ms` - 2 methods
- `self._svg_molecule_drag` - 4 methods

</details>

<details><summary>13 attributes, 0% isolated (0/4 methods)</summary>

- `self._copy_view_to_clipboard` - 2 methods
- `self._on_apply_fixed_crop_shortcut` - 3 methods
- `self._on_cancel_fixed_crop_shortcut` - 2 methods
- `self._on_clear_fixed_crop_shortcut` - 2 methods
- `self._set_active_view_ax` - 2 methods
- `self.enable_scale_bar` - 3 methods
- `self.set_angle_tool_enabled` - 2 methods
- `self.set_profile_tool_enabled` - 2 methods
- `self.set_show_acquisition_overlay` - 3 methods
- `self.set_show_angle_overlays` - 3 methods
- `self.set_show_molecules` - 3 methods
- `self.set_show_profile_overlays` - 3 methods
- `self.set_show_shortcut_hint` - 4 methods

</details>

<details><summary>6 attributes, 0% isolated (0/15 methods)</summary>

- `self._colorbar_orientation` - 8 methods
- `self._fit_to_canvas` - 7 methods
- `self._frame_fill_mode` - 6 methods
- `self._show_colorbar` - 11 methods
- `self._show_ticks` - 11 methods
- `self._show_title` - 11 methods

</details>

<details><summary>6 attributes, 0% isolated (0/12 methods)</summary>

- `self._profile_echo_artists` - 9 methods
- `self._profile_endpoint_labels` - 7 methods
- `self._profile_label` - 7 methods
- `self._profile_line` - 9 methods
- `self._profile_p0` - 9 methods
- `self._profile_p1` - 9 methods

</details>

### Spine attributes (touched by the most methods)

> These resist extraction and define what the class is *actually* about; everything else is a passenger.

| Attribute | Methods touching |
|---|---|
| `self.draw_idle` | 57 |
| `self.main_ax` | 55 |
| `self._redraw` | 49 |
| `self.push_undo_state` | 42 |
| `self._saved_profiles` | 30 |
| `self._ax_view_map` | 28 |
| `self.profile_pts` | 25 |
| `self._fixed_crop_template` | 22 |
| `self._notify_views_callback` | 22 |
| `self._angle_frames` | 20 |
| `self.views` | 20 |
| `self._outline_key` | 20 |
| `self.molecules` | 19 |
| `self.scale_bar_enabled` | 17 |
| `self._fixed_crop_template_bounds` | 17 |

