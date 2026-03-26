"""Context-menu helper for QLabel widgets that display QPixmaps."""
from __future__ import annotations

import types

from .._shared import QtCore, QtGui, QtWidgets
from .ppt_bridge import powerpoint_support_status, send_pixmap_to_ppt


class PPTContextMenuMixin:
    """Add live PowerPoint export actions to pixmap-backed QLabel widgets."""

    ppt_image_label: str | None = None

    def _resolve_ppt_label(self) -> str | None:
        label_text = getattr(self, "ppt_image_label", None)
        if label_text is not None:
            label_text = str(label_text).strip()
            if label_text:
                return label_text

        if hasattr(self, "toolTip"):
            tooltip = str(self.toolTip() or "").strip()
            if tooltip:
                return tooltip

        if hasattr(self, "objectName"):
            object_name = str(self.objectName() or "").strip()
            if object_name:
                return object_name
        return None

    def contextMenuEvent(self, event):
        menu = QtWidgets.QMenu(self)
        send_new_action = menu.addAction("Send to PowerPoint")
        send_current_action = menu.addAction("Send to Current Slide")

        supported, reason = powerpoint_support_status()
        if not supported:
            reason = reason or "PowerPoint export is unavailable."
            send_new_action.setEnabled(False)
            send_current_action.setEnabled(False)
            send_new_action.setToolTip(reason)
            send_current_action.setToolTip(reason)

        global_pos = QtGui.QCursor.pos()
        if event is not None:
            try:
                global_pos = event.globalPos()
            except Exception:
                pass

        chosen = menu.exec_(global_pos)
        if chosen == send_new_action:
            self._do_send_to_ppt(new_slide=True)
            return
        if chosen == send_current_action:
            self._do_send_to_ppt(new_slide=False)
            return

        try:
            super_context_menu = getattr(super(PPTContextMenuMixin, self), "contextMenuEvent", None)
        except TypeError:
            super_context_menu = None
        if callable(super_context_menu):
            super_context_menu(event)

    def _do_send_to_ppt(self, new_slide: bool):
        pixmap = self.pixmap() if hasattr(self, "pixmap") else None
        label_text = self._resolve_ppt_label()

        try:
            slide_number, shape_name = send_pixmap_to_ppt(
                pixmap,
                label=label_text,
                new_slide=bool(new_slide),
            )
        except ConnectionError:
            QtWidgets.QMessageBox.critical(
                self,
                "PowerPoint",
                "PowerPoint is not running. Please open a presentation first.",
            )
            return
        except ValueError:
            QtWidgets.QMessageBox.warning(self, "PowerPoint", "No image to send.")
            return
        except EnvironmentError as exc:
            QtWidgets.QMessageBox.critical(self, "PowerPoint", str(exc))
            return
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "PowerPoint", str(exc))
            return

        self._show_success(slide_number, shape_name)

    def _show_success(self, slide_number, shape_name):
        _ = shape_name
        target_rect = self.rect() if hasattr(self, "rect") else QtCore.QRect()
        QtWidgets.QToolTip.showText(
            QtGui.QCursor.pos(),
            f"Sent to slide {slide_number}",
            self,
            target_rect,
            2500,
        )

    @classmethod
    def install(cls, widget, label_text=None):
        if widget is None:
            return None

        if label_text is not None:
            widget.ppt_image_label = label_text

        if getattr(widget, "_ppt_context_menu_installed", False):
            return widget

        for name in ("contextMenuEvent", "_do_send_to_ppt", "_show_success", "_resolve_ppt_label"):
            setattr(widget, name, types.MethodType(getattr(cls, name), widget))

        widget._ppt_context_menu_installed = True
        widget.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        return widget
