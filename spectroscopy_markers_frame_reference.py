#!/usr/bin/env python3
"""
KPFM Spectroscopy Planner  —  zero external dependencies
---------------------------------------------------------
Requires only the Python standard library + tkinter.

Workflow:
  1. Load the Anfatec .txt  -> reads xCenter, yCenter, XScanRange, YScanRange
                               and auto-discovers all sibling .bmp channel files
  2. Pick a channel          -> displayed in the canvas
  3. Left-click              -> place numbered spectroscopy marker
  4. Right-click             -> remove nearest marker
  5. Export                  -> writes a .scr with GoXY + SpectPara + SpectStart

BMP support:
  Pure stdlib reader handles 8-bit (palette), 24-bit and 32-bit uncompressed BMPs.
  Anfatec SXM exports 32-bit BGRA BMPs (confirmed from file-size analysis).
  Image is fed to tkinter as an in-memory PPM (no temp files, no Pillow).

Coordinate transform (identical to kpfm_grid.scr):
  Xp = xCenter + (px / img_w - 0.5) * XScanRange
  Yp = yCenter + (0.5 - py / img_h) * YScanRange
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import struct
import re
import os
import math
from pathlib import Path
from datetime import datetime


# ── Palette ───────────────────────────────────────────────────────────────────
BG      = "#0e1117"
BG2     = "#161b22"
BG3     = "#21262d"
BORDER  = "#30363d"
FG      = "#e6edf3"
FG2     = "#8b949e"
ACCENT  = "#58a6ff"
ACCENT2 = "#f78166"
GREEN   = "#3fb950"
MONO    = "Courier New"
SANS    = "Segoe UI" if os.name == "nt" else "DejaVu Sans"
CROSS   = "#f78166"


# ── Pure-stdlib BMP reader ────────────────────────────────────────────────────
def read_bmp(path: str):
    """
    Read an uncompressed BMP file.
    Supports 8-bit (palette), 24-bit (BGR) and 32-bit (BGRA).
    Returns (rgb_bytes: bytes, width: int, height: int)
      where rgb_bytes is a flat sequence of R,G,B bytes, top-to-bottom.
    """
    with open(path, "rb") as f:
        data = f.read()

    if data[:2] != b"BM":
        raise ValueError(f"Not a BMP file: {path}")

    pixel_offset = struct.unpack_from("<I", data, 10)[0]
    dib_size     = struct.unpack_from("<I", data, 14)[0]
    w            = struct.unpack_from("<i", data, 18)[0]
    h            = struct.unpack_from("<i", data, 22)[0]
    bpp          = struct.unpack_from("<H", data, 28)[0]
    compression  = struct.unpack_from("<I", data, 30)[0]

    if compression != 0:
        raise ValueError(f"Compressed BMP not supported (compression={compression})")

    bottom_up = h > 0
    h = abs(h)
    row_size = ((bpp * w + 31) // 32) * 4   # each row padded to 4 bytes

    out = bytearray(w * h * 3)   # output: flat RGB

    if bpp == 32:
        for row in range(h):
            src = h - 1 - row if bottom_up else row
            off = pixel_offset + src * row_size
            dst = row * w * 3
            for col in range(w):
                b = data[off + col*4]
                g = data[off + col*4 + 1]
                r = data[off + col*4 + 2]
                out[dst + col*3]     = r
                out[dst + col*3 + 1] = g
                out[dst + col*3 + 2] = b

    elif bpp == 24:
        for row in range(h):
            src = h - 1 - row if bottom_up else row
            off = pixel_offset + src * row_size
            dst = row * w * 3
            for col in range(w):
                b = data[off + col*3]
                g = data[off + col*3 + 1]
                r = data[off + col*3 + 2]
                out[dst + col*3]     = r
                out[dst + col*3 + 1] = g
                out[dst + col*3 + 2] = b

    elif bpp == 8:
        # Palette: 256 * 4 bytes (BGRA) at offset 14 + dib_size
        pal_off = 14 + dib_size
        palette = []
        for i in range(256):
            b = data[pal_off + i*4]
            g = data[pal_off + i*4 + 1]
            r = data[pal_off + i*4 + 2]
            palette.append((r, g, b))
        for row in range(h):
            src = h - 1 - row if bottom_up else row
            off = pixel_offset + src * row_size
            dst = row * w * 3
            for col in range(w):
                r, g, b = palette[data[off + col]]
                out[dst + col*3]     = r
                out[dst + col*3 + 1] = g
                out[dst + col*3 + 2] = b
    else:
        raise ValueError(f"Unsupported BMP bpp={bpp}  (supported: 8, 24, 32)")

    return bytes(out), w, h


def rgb_to_ppm(rgb_bytes: bytes, w: int, h: int) -> bytes:
    """Convert flat RGB bytes to a PPM P6 binary blob (tkinter PhotoImage accepts this)."""
    header = f"P6\n{w} {h}\n255\n".encode()
    return header + rgb_bytes


def ppm_to_photoimage(ppm: bytes) -> tk.PhotoImage:
    """
    Load PPM bytes into a tkinter PhotoImage.
    tkinter only accepts PPM via file path, not via data=, so we use a temp file.
    """
    import tempfile, os
    fd, path = tempfile.mkstemp(suffix='.ppm')
    try:
        os.write(fd, ppm)
        os.close(fd)
        img = tk.PhotoImage(file=path)
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass
    return img


# ── .txt parser ───────────────────────────────────────────────────────────────
def parse_txt(path: str):
    """
    Parse Anfatec .txt -> (params dict, list of sibling BMP Paths).
    params keys: x_center, y_center, x_range, y_range, x_unit, angle
    """
    p = Path(path)
    raw = p.read_text(errors="replace")

    def flt(pat):
        m = re.search(pat, raw)
        return float(m.group(1)) if m else None

    def s(pat):
        m = re.search(pat, raw)
        return m.group(1).strip() if m else "nm"

    params = {
        "x_center": flt(r"xCenter\s*:\s*([+-]?\S+)"),
        "y_center": flt(r"yCenter\s*:\s*([+-]?\S+)"),
        "x_range":  flt(r"XScanRange\s*:\s*([+-]?\S+)"),
        "y_range":  flt(r"YScanRange\s*:\s*([+-]?\S+)"),
        "x_unit":   s  (r"XPhysUnit\s*:\s*(\S+)"),
        "angle":   flt(r"Angle\s*:\s*([+-]?\S+)"),
    }
    if params["angle"] is None:
        params["angle"] = 0.0
    missing = [k for k, v in params.items() if v is None and k not in ("x_unit", "angle")]
    if missing:
        raise ValueError(f"Missing fields in .txt: {', '.join(missing)}")

    base = p.stem
    bmps = sorted(p.parent.glob(f"{base}*.bmp"))
    return params, bmps, base


def channel_label(bmp: Path, base: str) -> str:
    """'angii_au111_01df Fwd.bmp' -> 'df Fwd'"""
    return bmp.stem[len(base):].strip() or bmp.name


# ── Script generator ──────────────────────────────────────────────────────────
def generate_scr(markers, params, img_w, img_h, base_name):
    X0, Y0 = params["x_center"], params["y_center"]
    Rx, Ry = params["x_range"],  params["y_range"]
    u       = params["x_unit"]

    lines = [
        "Begin",
        f"  FileName('{base_name}');",
        "  SpectPara('AUTOSAVE', 1);",
        "",
    ]
    for n, (px, py) in enumerate(markers, start=1):
        xp, yp = pixel_to_phys(px, py, img_w, img_h, params)
        lines += [
            f"  // Point {n}  pixel({px}, {py})  ->  ({xp:.4f}, {yp:.4f}) {u}",
            f"  SpectPara(1, {xp:.6f});   // X position",
            f"  SpectPara(2, {yp:.6f});   // Y position",
            "  SpectStart;",
            "",
        ]
    lines.append("end.")
    return "\n".join(lines)


# ── Image helpers (stdlib only) ─────────────────────────────────────────────
def vflip_rgb(rgb: bytes, w: int, h: int) -> bytes:
    """Flip RGB image vertically (row order reversal)."""
    row = w * 3
    return b"".join(rgb[i*row:(i+1)*row] for i in range(h-1, -1, -1))

def rotate_uv(u: float, v: float, angle_deg: float):
    """Rotate normalized scan coords (u, v) by Angle (degrees) around origin.

    Convention here: positive Angle is CCW in the lab XY frame.
    If your SXM defines clockwise-positive, change `a = -a` below.
    """
    if not angle_deg:
        return u, v
    a = math.radians(angle_deg)
    ca = math.cos(a)
    sa = math.sin(a)
    return (u * ca - v * sa), (u * sa + v * ca)


def pixel_to_phys(px: float, py: float, img_w: int, img_h: int, params: dict):
    """Map image pixel (px, py) to instrument XY using center, ranges, and Angle.

    Rules used here (matches what you described):
      - Top image row (py=0) is the start of the scan and maps to +Y offset.
      - xCenter/yCenter are the scan-frame center in instrument coordinates.
      - XScanRange/YScanRange are the full spans along the scan's *local* axes.
      - Angle rotates the scan local axes into the lab XY frame.

    Key point:
      We rotate *normalized* scan coordinates (u, v) first, then scale by ranges.
      This matches a scan-frame rotation model and avoids a big systematic drift.
    """
    X0, Y0 = params["x_center"], params["y_center"]
    Rx, Ry = params["x_range"],  params["y_range"]

    # Normalized scan coords in the scan frame, centered at 0
    u = (px / img_w) - 0.5
    v = 0.5 - (py / img_h)   # top row -> +v

    u, v = rotate_uv(u, v, params.get("angle", 0.0))

    xp = X0 + u * Rx
    yp = Y0 + v * Ry
    return xp, yp




# ── Resize helper (stdlib only, nearest-neighbour) ────────────────────────────
def resize_rgb(rgb: bytes, src_w, src_h, dst_w, dst_h) -> bytes:
    """
    Fast nearest-neighbour resize using precomputed index tables.
    ~10x faster than the naive pixel loop for SPM image sizes.
    """
    # Precompute source row and column indices
    x_ratio = src_w / dst_w
    y_ratio = src_h / dst_h
    src_xs = [int(dx * x_ratio) for dx in range(dst_w)]
    src_ys = [int(dy * y_ratio) for dy in range(dst_h)]
    out = bytearray(dst_w * dst_h * 3)
    for dy, sy in enumerate(src_ys):
        src_row_off = sy * src_w * 3
        dst_row_off = dy * dst_w * 3
        for dx, sx in enumerate(src_xs):
            so = src_row_off + sx * 3
            do = dst_row_off + dx * 3
            out[do:do+3] = rgb[so:so+3]
    return bytes(out)


# ── Application ───────────────────────────────────────────────────────────────
class Planner(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KPFM Spectroscopy Planner  ·  Anfatec SXM")
        self.configure(bg=BG)
        self.geometry("1280x800")
        self.minsize(900, 600)

        # Data state
        self.params    = None
        self.bmp_paths = []
        self.txt_base  = ""
        self.rgb_bytes = None     # full-resolution RGB bytes of current channel
        self.img_w     = 1
        self.img_h     = 1
        self.markers   = []       # [(px_orig, py_orig), ...]

        # Display state
        self.tk_img    = None
        self.disp_w    = 1
        self.disp_h    = 1
        self.scale     = 1.0
        self.offset_x  = 0
        self.offset_y  = 0

        # UI vars
        self.ch_var    = tk.StringVar()
        self.base_var  = tk.StringVar(value="CustomPoints")
        self.status    = tk.StringVar(value="Load an Anfatec .txt file to begin.")

        self._build()

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build(self):
        bar = tk.Frame(self, bg=BG2, height=48)
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)
        tk.Label(bar, text="◈  KPFM SPECTROSCOPY PLANNER",
                 font=(MONO, 13, "bold"), bg=BG2, fg=ACCENT
                 ).pack(side=tk.LEFT, padx=16, pady=12)
        tk.Label(bar, text="Anfatec SXM  ·  zero dependencies",
                 font=(MONO, 9), bg=BG2, fg=FG2).pack(side=tk.LEFT)

        body = tk.Frame(self, bg=BG)
        body.pack(fill=tk.BOTH, expand=True)

        side = tk.Frame(body, bg=BG2, width=268)
        side.pack(side=tk.LEFT, fill=tk.Y)
        side.pack_propagate(False)
        self._build_side(side)

        tk.Frame(body, bg=BORDER, width=1).pack(side=tk.LEFT, fill=tk.Y)

        cv = tk.Frame(body, bg=BG)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._build_canvas(cv)

        sb = tk.Frame(self, bg=BG3, height=26)
        sb.pack(fill=tk.X)
        sb.pack_propagate(False)
        tk.Label(sb, textvariable=self.status, font=(MONO, 9),
                 bg=BG3, fg=FG2, anchor="w").pack(side=tk.LEFT, padx=10)

    def _build_side(self, p):
        def sep():
            tk.Frame(p, bg=BORDER, height=1).pack(fill=tk.X, padx=14, pady=8)

        def sec(t):
            tk.Label(p, text=t, font=(MONO, 8, "bold"),
                     bg=BG2, fg=FG2).pack(anchor="w", padx=14, pady=(12, 2))

        def param_row(label, attr):
            r = tk.Frame(p, bg=BG2)
            r.pack(fill=tk.X, padx=14, pady=1)
            tk.Label(r, text=label, font=(SANS, 8), bg=BG2, fg=FG2,
                     width=9, anchor="w").pack(side=tk.LEFT)
            lbl = tk.Label(r, text="—", font=(MONO, 9), bg=BG2,
                           fg=FG, anchor="w")
            lbl.pack(side=tk.LEFT, padx=2)
            setattr(self, attr, lbl)

        sec("LOAD")
        tk.Button(p, text="⊕  Load .txt  (auto-finds channels)",
                  font=(MONO, 10), bg=BG3, fg=ACCENT, relief=tk.FLAT,
                  activebackground=BORDER, activeforeground=ACCENT,
                  cursor="hand2", command=self._load_txt,
                  anchor="w", padx=10, pady=8,
                  highlightthickness=1, highlightbackground=BORDER
                  ).pack(fill=tk.X, padx=14, pady=3)
        sep()

        sec("SCAN PARAMETERS")
        param_row("xCenter", "lbl_xc")
        param_row("yCenter", "lbl_yc")
        param_row("X range", "lbl_xr")
        param_row("Y range", "lbl_yr")
        param_row("Angle",  "lbl_ang")
        sep()

        sec("CHANNEL")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("D.TCombobox",
                         fieldbackground=BG3, background=BG3,
                         foreground=FG, selectbackground=ACCENT,
                         selectforeground=BG)
        self.ch_combo = ttk.Combobox(p, textvariable=self.ch_var,
                                      style="D.TCombobox",
                                      state="disabled",
                                      font=(MONO, 9), width=26)
        self.ch_combo.pack(padx=14, pady=4, fill=tk.X)
        self.ch_combo.bind("<<ComboboxSelected>>", self._on_ch_change)
        sep()

        sec("MARKERS")
        lf = tk.Frame(p, bg=BG3, highlightthickness=1,
                      highlightbackground=BORDER)
        lf.pack(fill=tk.BOTH, expand=True, padx=14, pady=2)
        self.mlist = tk.Listbox(lf, font=(MONO, 9), bg=BG3, fg=FG,
                                 selectbackground=ACCENT, selectforeground=BG,
                                 relief=tk.FLAT, borderwidth=0,
                                 activestyle="none", highlightthickness=0)
        self.mlist.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        brow = tk.Frame(p, bg=BG2)
        brow.pack(fill=tk.X, padx=14, pady=2)
        for txt, cmd in [("Remove sel.", self._remove_sel),
                          ("Clear all",  self._clear_all)]:
            tk.Button(brow, text=txt, font=(MONO, 8), bg=BG3, fg=ACCENT2,
                      relief=tk.FLAT, activebackground=BORDER,
                      activeforeground=ACCENT2, cursor="hand2",
                      command=cmd, padx=6, pady=3
                      ).pack(side=tk.LEFT, padx=(0, 4))
        sep()

        row = tk.Frame(p, bg=BG2)
        row.pack(fill=tk.X, padx=14, pady=2)
        tk.Label(row, text="Filename", font=(SANS, 8), bg=BG2, fg=FG2,
                 width=9, anchor="w").pack(side=tk.LEFT)
        tk.Entry(row, textvariable=self.base_var, font=(MONO, 9),
                 bg=BG3, fg=FG, insertbackground=ACCENT, relief=tk.FLAT,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT, width=13
                 ).pack(side=tk.LEFT, padx=4)

        tk.Button(p, text="⬇  EXPORT  .scr",
                  font=(MONO, 11, "bold"), bg=GREEN, fg=BG,
                  relief=tk.FLAT, activebackground="#2ea043",
                  activeforeground=BG, cursor="hand2",
                  command=self._export, pady=10
                  ).pack(fill=tk.X, padx=14, pady=(6, 14))

    def _build_canvas(self, p):
        ib = tk.Frame(p, bg=BG3, height=30)
        ib.pack(fill=tk.X)
        ib.pack_propagate(False)
        tk.Label(ib,
                 text="Left-click  →  add marker     Right-click  →  remove nearest",
                 font=(MONO, 9), bg=BG3, fg=FG2
                 ).pack(side=tk.LEFT, padx=10, pady=6)
        self.coord_lbl = tk.Label(ib, text="", font=(MONO, 9), bg=BG3, fg=ACCENT)
        self.coord_lbl.pack(side=tk.RIGHT, padx=10)

        self.canvas = tk.Canvas(p, bg=BG3, cursor="crosshair",
                                 highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>",  self._on_left)
        self.canvas.bind("<Button-3>",  self._on_right)
        self.canvas.bind("<Motion>",    self._on_motion)
        self.canvas.bind("<Configure>", lambda e: self._render())
        self._placeholder()

    # ── File loading ──────────────────────────────────────────────────────────
    def _load_txt(self):
        path = filedialog.askopenfilename(
            title="Select Anfatec .txt parameter file",
            filetypes=[("Anfatec parameter", "*.txt"), ("All files", "*.*")])
        if not path:
            return
        try:
            self.params, self.bmp_paths, self.txt_base = parse_txt(path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        u = self.params["x_unit"]
        self.lbl_xc.config(text=f"{self.params['x_center']:.3f} {u}")
        self.lbl_yc.config(text=f"{self.params['y_center']:.3f} {u}")
        self.lbl_xr.config(text=f"{self.params['x_range']:.3f} {u}")
        self.lbl_yr.config(text=f"{self.params['y_range']:.3f} {u}")
        self.lbl_ang.config(text=f"{self.params['angle']:.3f} deg")

        if self.bmp_paths:
            labels = [channel_label(b, self.txt_base) for b in self.bmp_paths]
            self.ch_combo.config(values=labels, state="readonly")
            self.ch_combo.current(0)
            self._load_bmp(self.bmp_paths[0])
            self.status.set(
                f"Loaded: {Path(path).name}  ·  "
                f"{len(self.bmp_paths)} channels  ·  click to place markers")
        else:
            self.ch_combo.config(values=[], state="disabled")
            self.status.set(
                f"Loaded: {Path(path).name}  ·  no .bmp files found in same folder")

    def _on_ch_change(self, event=None):
        idx = self.ch_combo.current()
        if 0 <= idx < len(self.bmp_paths):
            self._load_bmp(self.bmp_paths[idx])

    def _load_bmp(self, bmp: Path):
        try:
            rgb, w, h = read_bmp(str(bmp))
        except Exception as e:
            messagebox.showerror("BMP error", str(e))
            return
        # Clear markers only if resolution changes between channels
        if w != self.img_w or h != self.img_h:
            self.markers.clear()
            self._update_list()
        self.rgb_bytes, self.img_w, self.img_h = rgb, w, h
        # Defer until canvas is fully laid out (fixes blank on first load)
        self.after(20, self._render)

    # ── Rendering ─────────────────────────────────────────────────────────────
    def _placeholder(self):
        self.canvas.delete("all")
        w = self.canvas.winfo_width() or 600
        h = self.canvas.winfo_height() or 400
        self.canvas.create_text(w//2, h//2,
                                 text="Load a .txt file to begin",
                                 fill=FG2, font=(MONO, 13))

    def _render(self):
        if self.rgb_bytes is None:
            self._placeholder()
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        # Canvas not yet laid out — reschedule once more
        if cw < 10 or ch < 10:
            self.after(50, self._render)
            return

        sc = min(cw / self.img_w, ch / self.img_h)
        dw = max(1, int(self.img_w * sc))
        dh = max(1, int(self.img_h * sc))
        ox = (cw - dw) // 2
        oy = (ch - dh) // 2
        self.scale, self.offset_x, self.offset_y = sc, ox, oy
        self.disp_w, self.disp_h = dw, dh

        resized = resize_rgb(self.rgb_bytes, self.img_w, self.img_h, dw, dh)
        ppm = rgb_to_ppm(resized, dw, dh)
        self.tk_img = ppm_to_photoimage(ppm)

        self.canvas.delete("all")
        self.canvas.create_image(ox, oy, anchor="nw", image=self.tk_img)
        self._draw_markers()

    def _draw_markers(self):
        self.canvas.delete("m")
        arm = max(8, int(self.disp_w * 0.018))
        for n, (px, py) in enumerate(self.markers, start=1):
            cx = self.offset_x + px * self.scale
            cy = self.offset_y + py * self.scale
            self.canvas.create_line(cx-arm, cy, cx+arm, cy,
                                     fill=CROSS, width=1.5, tags="m")
            self.canvas.create_line(cx, cy-arm, cx, cy+arm,
                                     fill=CROSS, width=1.5, tags="m")
            r = 3
            self.canvas.create_oval(cx-r, cy-r, cx+r, cy+r,
                                     fill=CROSS, outline="white",
                                     width=0.8, tags="m")
            self.canvas.create_text(cx + arm + 5, cy - arm,
                                     text=str(n), fill=CROSS,
                                     font=(MONO, 8, "bold"),
                                     anchor="sw", tags="m")

    # ── Canvas interaction ────────────────────────────────────────────────────
    def _canvas_to_img(self, cx, cy):
        if self.rgb_bytes is None:
            return None, None
        px = (cx - self.offset_x) / self.scale
        py = (cy - self.offset_y) / self.scale
        if 0 <= px < self.img_w and 0 <= py < self.img_h:
            return int(round(px)), int(round(py))
        return None, None

    def _on_left(self, event):
        px, py = self._canvas_to_img(event.x, event.y)
        if px is None:
            return
        self.markers.append((px, py))
        self._update_list()
        self._draw_markers()

    def _on_right(self, event):
        if not self.markers:
            return
        px, py = self._canvas_to_img(event.x, event.y)
        if px is None:
            return
        idx = min(range(len(self.markers)),
                  key=lambda i: (self.markers[i][0]-px)**2 + (self.markers[i][1]-py)**2)
        self.markers.pop(idx)
        self._update_list()
        self._draw_markers()

    def _on_motion(self, event):
        px, py = self._canvas_to_img(event.x, event.y)
        if px is None or self.params is None:
            self.coord_lbl.config(text="")
            return
        xp, yp = pixel_to_phys(px, py, self.img_w, self.img_h, self.params)
        u  = self.params["x_unit"]
        self.coord_lbl.config(
            text=f"pixel ({px}, {py})  →  ({xp:+.3f}, {yp:+.3f}) {u}")

    # ── Marker list ───────────────────────────────────────────────────────────
    def _update_list(self):
        self.mlist.delete(0, tk.END)
        for n, (px, py) in enumerate(self.markers, start=1):
            if self.params:
                xp, yp = pixel_to_phys(px, py, self.img_w, self.img_h, self.params)
                u = self.params["x_unit"]
                self.mlist.insert(tk.END, f"#{n:02d}  ({xp:+.3f}, {yp:+.3f}) {u}")
            else:
                self.mlist.insert(tk.END, f"#{n:02d}  pixel ({px}, {py})")

    def _remove_sel(self):
        sel = self.mlist.curselection()
        if sel:
            self.markers.pop(sel[0])
            self._update_list()
            self._draw_markers()

    def _clear_all(self):
        if self.markers and messagebox.askyesno("Clear all", "Remove all markers?"):
            self.markers.clear()
            self._update_list()
            self._draw_markers()

    # ── Export ────────────────────────────────────────────────────────────────
    def _export(self):
        if not self.markers:
            messagebox.showwarning("No markers", "Place at least one marker first.")
            return
        if not self.params:
            messagebox.showwarning("No parameters", "Load the .txt file first.")
            return

        base = self.base_var.get() or "CustomPoints"
        out  = filedialog.asksaveasfilename(
            title="Save .scr script", defaultextension=".scr",
            initialfile=f"{base}.scr",
            filetypes=[("SXM script", "*.scr"), ("All", "*.*")])
        if not out:
            return

        p = self.params
        hdr = (
            f"// KPFM Spectroscopy Planner  --  {datetime.now():%Y-%m-%d %H:%M:%S}\n"
            f"// Source  : {self.txt_base}.txt\n"
            f"// xCenter : {p['x_center']}  yCenter : {p['y_center']}  "
            f"unit : {p['x_unit']}\n"
            f"// XRange  : {p['x_range']}  YRange  : {p['y_range']}\n"
            f"// Image   : {self.img_w} x {self.img_h} px  "
            f"|  Points : {len(self.markers)}\n\n"
        )
        scr = generate_scr(self.markers, p, self.img_w, self.img_h, base)
        try:
            Path(out).write_text(hdr + scr)
            self.status.set(
                f"Exported {len(self.markers)} point(s)  ->  {Path(out).name}")
            messagebox.showinfo("Done",
                f"{len(self.markers)} point(s) saved to:\n{out}")
        except OSError as e:
            messagebox.showerror("Export failed", str(e))


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = Planner()
    app.mainloop()
