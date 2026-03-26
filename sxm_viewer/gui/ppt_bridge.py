"""PowerPoint COM bridge for live image export on Windows."""
from __future__ import annotations

import logging
import os
import sys
import tempfile


LOGGER = logging.getLogger(__name__)
PPT_LAYOUT_BLANK = 12
MSO_TEXT_ORIENTATION_HORIZONTAL = 1

pythoncom = None
win32com = None
_WIN32_IMPORT_ERROR = None
HAS_WIN32 = False

if sys.platform == "win32":
    try:
        import pythoncom as _pythoncom
        import win32com.client as _win32com_client
    except Exception as exc:  # pragma: no cover - depends on local Windows setup
        _WIN32_IMPORT_ERROR = exc
    else:  # pragma: no branch
        pythoncom = _pythoncom
        win32com = _win32com_client
        HAS_WIN32 = True


def powerpoint_support_status() -> tuple[bool, str | None]:
    """Return whether the live PowerPoint bridge can run in this environment."""
    if sys.platform != "win32":
        return False, "PowerPoint export is only available on Windows."
    if not HAS_WIN32:
        return (
            False,
            "pywin32 is required for PowerPoint export. Install it with 'pip install pywin32'.",
        )
    return True, None


class PowerPointBridge:
    """Reusable COM bridge to an already-running PowerPoint instance."""

    def __init__(self):
        self._app = None

    def _connect(self) -> bool:
        supported, message = powerpoint_support_status()
        if not supported:
            raise EnvironmentError(message or "PowerPoint export is unavailable.")

        try:  # pragma: no cover - no-op on already initialized threads
            pythoncom.CoInitialize()
        except Exception:
            pass

        if self._app is not None:
            try:
                if int(self._app.Presentations.Count) > 0:
                    _ = self._app.ActivePresentation.Name
                    return True
            except Exception:
                self._app = None

        try:
            app = win32com.GetActiveObject("PowerPoint.Application")
        except Exception:
            self._app = None
            return False

        try:
            if int(app.Presentations.Count) < 1:
                self._app = None
                return False
            _ = app.ActivePresentation.Name
        except Exception:
            self._app = None
            return False

        self._app = app
        return True

    def _presentation(self):
        if not self._connect():
            raise ConnectionError(
                "PowerPoint is not running, or there is no open presentation."
            )
        try:
            return self._app.ActivePresentation
        except Exception as exc:
            raise ConnectionError(
                "PowerPoint is running, but no presentation is currently active."
            ) from exc

    def _resolve_slide(self, presentation, *, new_slide: bool, slide_index: int | None):
        slides = presentation.Slides
        slide_count = int(slides.Count)

        if slide_index is not None:
            try:
                index = int(slide_index)
            except Exception as exc:
                raise ValueError(f"Invalid slide index: {slide_index!r}") from exc
            if index < 1 or index > slide_count:
                raise ValueError(
                    f"Slide index {index} is out of range. Presentation has {slide_count} slide(s)."
                )
            return slides.Item(index)

        if new_slide:
            return slides.Add(slide_count + 1, PPT_LAYOUT_BLANK)

        try:
            active_window = self._app.ActiveWindow
            view = active_window.View
            slide = view.Slide
            if slide is None:
                raise RuntimeError("No active slide view.")
            return slide
        except Exception as exc:
            raise ConnectionError(
                "PowerPoint does not have an active slide view. Activate a slide and try again."
            ) from exc

    def _fit_image_box(
        self,
        *,
        left: float,
        top: float,
        width: float,
        height: float,
        image_size: tuple[int, int] | None,
        preserve_aspect: bool,
    ) -> tuple[float, float, float, float]:
        box_left = float(left)
        box_top = float(top)
        box_width = max(float(width), 1.0)
        box_height = max(float(height), 1.0)

        if not preserve_aspect or not image_size:
            return box_left, box_top, box_width, box_height

        try:
            px_width = max(float(image_size[0]), 1.0)
            px_height = max(float(image_size[1]), 1.0)
        except Exception:
            return box_left, box_top, box_width, box_height

        aspect = px_width / px_height
        fitted_width = box_width
        fitted_height = fitted_width / aspect
        if fitted_height > box_height:
            fitted_height = box_height
            fitted_width = fitted_height * aspect

        fitted_left = box_left + (box_width - fitted_width) * 0.5
        fitted_top = box_top + (box_height - fitted_height) * 0.5
        return fitted_left, fitted_top, fitted_width, fitted_height

    def send_image(
        self,
        image_path,
        *,
        new_slide: bool = True,
        slide_index: int | None = None,
        left: float = 50,
        top: float = 50,
        width: float = 600,
        height: float = 450,
        label: str | None = None,
        image_size: tuple[int, int] | None = None,
        preserve_aspect: bool = True,
    ) -> tuple[int, str]:
        """Insert an image file into the active PowerPoint presentation."""
        if not image_path:
            raise ValueError("Image path is required.")

        image_path = os.path.abspath(os.fspath(image_path))
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Image path does not exist: {image_path}")

        presentation = self._presentation()
        slide = self._resolve_slide(
            presentation,
            new_slide=bool(new_slide),
            slide_index=slide_index,
        )
        shape_left, shape_top, shape_width, shape_height = self._fit_image_box(
            left=left,
            top=top,
            width=width,
            height=height,
            image_size=image_size,
            preserve_aspect=bool(preserve_aspect),
        )

        shape = slide.Shapes.AddPicture(
            FileName=image_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=shape_left,
            Top=shape_top,
            Width=shape_width,
            Height=shape_height,
        )

        label_text = str(label).strip() if label is not None else ""
        if label_text:
            text_box = slide.Shapes.AddTextbox(
                MSO_TEXT_ORIENTATION_HORIZONTAL,
                shape_left,
                shape_top + shape_height + 6.0,
                shape_width,
                24.0,
            )
            text_range = text_box.TextFrame.TextRange
            text_range.Text = label_text
            try:
                text_range.Font.Size = 12
            except Exception:
                pass
            try:
                text_box.Line.Visible = False
                text_box.Fill.Visible = False
            except Exception:
                pass

        return int(slide.SlideIndex), str(shape.Name)


_bridge = PowerPointBridge()


def send_pixmap_to_ppt(
    pixmap,
    label: str | None = None,
    **kwargs,
) -> tuple[int, str]:
    """Encode a QPixmap to PNG and send it to a live PowerPoint presentation."""
    from PyQt5 import QtCore

    if pixmap is None or not hasattr(pixmap, "isNull") or pixmap.isNull():
        raise ValueError("No image to send.")

    buffer = QtCore.QBuffer()
    if not buffer.open(QtCore.QIODevice.WriteOnly):
        raise OSError("Unable to open an in-memory image buffer.")

    try:
        if not pixmap.save(buffer, "PNG"):
            raise ValueError("Unable to encode the image as PNG.")
        png_bytes = bytes(buffer.data())
    finally:
        buffer.close()

    if not png_bytes:
        raise ValueError("Unable to encode the image as PNG.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix="sxm_viewer_ppt_",
            suffix=".png",
            delete=False,
        ) as handle:
            handle.write(png_bytes)
            tmp_path = handle.name
        image_size = kwargs.pop("image_size", (int(pixmap.width()), int(pixmap.height())))
        return _bridge.send_image(tmp_path, label=label, image_size=image_size, **kwargs)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError as exc:
                LOGGER.warning(
                    "Failed to delete temporary PowerPoint image '%s': %s",
                    tmp_path,
                    exc,
                )
