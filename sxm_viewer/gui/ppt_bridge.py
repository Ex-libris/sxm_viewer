"""PowerPoint COM bridge for live image export on Windows."""
from __future__ import annotations

import logging
import os
import re
import sys
import tempfile


LOGGER = logging.getLogger(__name__)
PPT_LAYOUT_BLANK = 12
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_GRAPHIC = 28  # MsoShapeType.msoGraphic (inserted SVG)

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
        convert_to_shapes: bool = False,
        fallback_path_factory=None,
    ) -> tuple[int, str]:
        """Insert an image file into the active PowerPoint presentation.

        With ``convert_to_shapes`` (SVG input), the inserted graphic is run
        through PowerPoint's own "Convert to Shape" so it lands as one native,
        ungroupable group of editable shapes instead of an opaque SVG icon.
        """
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

        def _add_picture(path):
            return slide.Shapes.AddPicture(
                FileName=path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=shape_left,
                Top=shape_top,
                Width=shape_width,
                Height=shape_height,
            )

        try:
            shape = _add_picture(image_path)
        except Exception as exc:
            # Office builds without SVG support reject the file; retry with a
            # raster rendering on the same slide (no stray empty slide).
            if fallback_path_factory is None:
                raise
            LOGGER.info("PowerPoint rejected %s, falling back to raster: %s", image_path, exc)
            fallback_path = fallback_path_factory()
            if not fallback_path:
                raise
            shape = _add_picture(fallback_path)
            convert_to_shapes = False
        if convert_to_shapes:
            shape = self._convert_graphic_to_shapes(slide, shape)

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

    def _convert_graphic_to_shapes(self, slide, shape):
        """Run "Convert to Shape" on an inserted SVG graphic; best effort.

        ``Shape.Ungroup`` refuses SVG graphics over COM ("only accessed for a
        group"), so the ribbon command is used instead, which needs the shape
        selected in a visible slide view. Any failure leaves the (still
        vector) SVG graphic in place.
        """
        try:
            if int(shape.Type) != MSO_GRAPHIC:
                return shape
            window = self._app.ActiveWindow
            if int(window.View.Slide.SlideIndex) != int(slide.SlideIndex):
                window.View.GotoSlide(int(slide.SlideIndex))
            shape.Select()
            command_bars = self._app.CommandBars
            if not command_bars.GetEnabledMso("SVGEdit"):
                return shape
            command_bars.ExecuteMso("SVGEdit")
            selection = window.Selection.ShapeRange
            if int(selection.Count) >= 1:
                return selection.Item(1)
        except Exception as exc:  # pragma: no cover - depends on Office build
            LOGGER.info("PowerPoint 'Convert to Shape' unavailable: %s", exc)
        return shape


_bridge = PowerPointBridge()

_SVG_LENGTH_UNITS_TO_PT = {"pt": 1.0, "px": 0.75, "in": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4, "": 1.0}


def _svg_size_pt(svg_bytes: bytes) -> tuple[float, float] | None:
    """Read the root ``width``/``height`` of an SVG document, in points."""
    head = svg_bytes[:4096].decode("utf-8", errors="ignore")
    match = re.search(r"<svg\b[^>]*>", head, flags=re.S)
    if not match:
        return None
    size = []
    for attr in ("width", "height"):
        found = re.search(rf'\b{attr}="([0-9.]+)\s*([a-z]*)"', match.group(0))
        if not found:
            return None
        scale = _SVG_LENGTH_UNITS_TO_PT.get(found.group(2))
        if scale is None:
            return None
        size.append(float(found.group(1)) * scale)
    return size[0], size[1]


def send_svg_to_ppt(
    svg_bytes: bytes,
    label: str | None = None,
    *,
    convert_to_shapes: bool = True,
    **kwargs,
) -> tuple[int, str]:
    """Send SVG bytes to a live PowerPoint presentation as vector content."""
    if not svg_bytes:
        raise ValueError("No image to send.")
    tmp_paths = []

    def _write_temp(data: bytes, suffix: str) -> str:
        with tempfile.NamedTemporaryFile(
            prefix="sxm_viewer_ppt_",
            suffix=suffix,
            delete=False,
        ) as handle:
            handle.write(data)
            tmp_paths.append(handle.name)
            return handle.name

    def _raster_fallback() -> str | None:
        png_bytes = _rasterize_svg(svg_bytes, kwargs.get("image_size"))
        return _write_temp(png_bytes, ".png") if png_bytes else None

    try:
        if "image_size" not in kwargs:
            kwargs["image_size"] = _svg_size_pt(svg_bytes)
        return _bridge.send_image(
            _write_temp(svg_bytes, ".svg"),
            label=label,
            convert_to_shapes=convert_to_shapes,
            fallback_path_factory=_raster_fallback,
            **kwargs,
        )
    finally:
        for tmp_path in tmp_paths:
            if not os.path.exists(tmp_path):
                continue
            try:
                os.unlink(tmp_path)
            except OSError as exc:
                LOGGER.warning(
                    "Failed to delete temporary PowerPoint image '%s': %s",
                    tmp_path,
                    exc,
                )


def send_layered_svg_to_ppt(
    svg_bytes: bytes,
    molecule_names: dict | None = None,
    label: str | None = None,
    **kwargs,
) -> tuple[int, str]:
    """Send a molecule-tagged figure SVG so each molecule lands as its own
    nested group (Bonds/Atoms/Labels inside), all within one figure group.

    PowerPoint's "Convert to Shape" flattens SVG ``<g>`` structure, so the
    base figure and every molecule role are converted from separate SVGs
    (same canvas, same box, so they overlay exactly) and regrouped here.
    """
    from . import vector_export

    base_svg, molecules = vector_export.split_for_powerpoint(svg_bytes, molecule_names)
    if not molecules:
        return send_svg_to_ppt(svg_bytes, label=label, **kwargs)

    if "image_size" not in kwargs:
        kwargs["image_size"] = _svg_size_pt(svg_bytes)
    slide_number, base_name = send_svg_to_ppt(base_svg, label=label, **kwargs)
    part_kwargs = dict(kwargs, new_slide=False, slide_index=slide_number)
    slide = _bridge._presentation().Slides.Item(int(slide_number))
    # Shapes are tracked by Id and grouped by index: display names repeat
    # ("Bonds"/"Atoms" per molecule), so Shapes.Range(names) is ambiguous.
    # A group's name survives being nested, so each is named as it's built
    # (ParentGroup only ever reports the outermost group, so names can't be
    # fixed up afterwards).

    def _named(shape, display_name: str) -> int:
        try:
            shape.Name = display_name
        except Exception:
            pass
        return int(shape.Id)

    def _group(shape_ids, display_name: str) -> int:
        if len(shape_ids) == 1:
            return _named(_shape_by_id(slide, shape_ids[0]), display_name)
        indices = _shape_indices(slide, shape_ids)
        return _named(slide.Shapes.Range(indices).Group(), display_name)

    try:
        members = [_named(slide.Shapes(base_name), "Image")]
        for mol_name, parts in molecules:
            role_ids = []
            for role_title, role_svg in parts:
                before = {int(shape.Id) for shape in slide.Shapes}
                send_svg_to_ppt(role_svg, label=None, **part_kwargs)
                added = [int(shape.Id) for shape in slide.Shapes if int(shape.Id) not in before]
                if added:
                    role_ids.append(_group(added, role_title))
            if role_ids:
                members.append(_group(role_ids, mol_name))
        figure = _shape_by_id(slide, _group(members, "SXM figure"))
        return slide_number, str(figure.Name)
    except Exception as exc:  # pragma: no cover - depends on Office build
        LOGGER.warning("PowerPoint molecule regrouping failed: %s", exc)
        return slide_number, base_name


def _shape_by_id(slide, shape_id: int):
    for shape in slide.Shapes:
        if int(shape.Id) == int(shape_id):
            return shape
    raise KeyError(f"Shape id {shape_id} not found on slide.")


def _shape_indices(slide, shape_ids) -> list[int]:
    wanted = {int(shape_id) for shape_id in shape_ids}
    return [
        idx
        for idx in range(1, int(slide.Shapes.Count) + 1)
        if int(slide.Shapes.Item(idx).Id) in wanted
    ]


def send_rendered_to_ppt(payload, label: str | None = None, **kwargs) -> tuple[int, str]:
    """Send a vector payload (the default), SVG bytes, or a QPixmap (raster)."""
    if isinstance(payload, dict):
        return send_layered_svg_to_ppt(
            payload.get("svg") or b"",
            payload.get("molecule_names"),
            label=label,
            **kwargs,
        )
    if isinstance(payload, (bytes, bytearray)):
        return send_svg_to_ppt(bytes(payload), label=label, **kwargs)
    return send_pixmap_to_ppt(payload, label=label, **kwargs)


def _rasterize_svg(svg_bytes: bytes, size_pt, dpi: float = 300.0) -> bytes | None:
    """Render SVG bytes to PNG (for Office builds that cannot insert SVG)."""
    from PyQt5 import QtCore, QtGui
    from PyQt5.QtSvg import QSvgRenderer

    renderer = QSvgRenderer(QtCore.QByteArray(svg_bytes))
    if not renderer.isValid():
        return None
    if size_pt:
        width_px = max(1, int(round(float(size_pt[0]) * dpi / 72.0)))
        height_px = max(1, int(round(float(size_pt[1]) * dpi / 72.0)))
    else:
        default = renderer.defaultSize()
        width_px, height_px = max(1, default.width()), max(1, default.height())
    image = QtGui.QImage(width_px, height_px, QtGui.QImage.Format_ARGB32)
    image.fill(QtCore.Qt.white)
    painter = QtGui.QPainter(image)
    try:
        renderer.render(painter)
    finally:
        painter.end()
    buffer = QtCore.QBuffer()
    buffer.open(QtCore.QIODevice.WriteOnly)
    try:
        if not image.save(buffer, "PNG"):
            return None
        return bytes(buffer.data())
    finally:
        buffer.close()


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
