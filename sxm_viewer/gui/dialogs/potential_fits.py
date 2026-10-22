"""Analytic potential fits (Lennard-Jones, Morse) for spectroscopy curves.

These fit a smooth model V(x) to an experimental trace so the model can be
overlaid on its curve and, optionally, differentiated analytically to give the
*force of the fit* (F = -dV/dx) - a clean alternative to numerically
deconvolving a noisy experimental curve.

Both potentials diverge on one side (LJ ~ (x-x0)^-12 near its singularity;
Morse ~ exp(-a x) as x -> -inf), so every model here is parameterised with the
singularity/steep wall pinned strictly *outside* the data range (x0 < min(x)),
and evaluation helpers clip to the fitted data span. Nothing is ever evaluated
at the pole, so the "explodes at z=0" failure mode cannot occur on screen.

Robust mode (the default) targets real qPlus/CO-tip Δf(z) curves, whose
short-range side is often not a clean wall: the CO bends, so Δf shoots up and
then turns back down at the closest approach, sometimes dipping below the real
minimum. A plain least-squares fit spends its freedom chasing those few large
residuals and loses the long range and the well. Robust mode therefore
(1) locates the attractive well coming in from the far side on a smoothed copy
(the first minimum met, not the global one), (2) keeps the repulsive branch
only up to ``wall_frac`` x the well depth above the minimum and never past its
turning point, (3) fits with a soft-L1 loss scaled to the noise so isolated
spikes cannot dominate, and (4) multi-starts the optimiser, keeping the best.
"""
from __future__ import annotations

import numpy as np

try:  # SciPy is an optional dependency elsewhere in the app; degrade cleanly.
    from scipy.optimize import curve_fit as _curve_fit
except Exception:  # pragma: no cover - exercised only where SciPy is absent
    _curve_fit = None
try:
    from scipy.signal import savgol_filter as _savgol
except Exception:  # pragma: no cover
    _savgol = None


MODELS = ("Lennard-Jones", "Morse")

# Smallest fraction of the data span the pole is allowed to approach the data.
_POLE_MARGIN_FRAC = 1e-3

# Robust mode: how far up the repulsive wall (in units of the well depth above
# the minimum) samples are still fitted.
DEFAULT_WALL_FRAC = 0.5


def scipy_available() -> bool:
    return _curve_fit is not None


# ---------------------------------------------------------------------------
# Model evaluation (safe: never divides by ~0, never overflows to inf/nan)
# ---------------------------------------------------------------------------
def _lj_value(x, eps, rm, x0, c):
    u = np.asarray(x, dtype=float) - x0
    # u is > 0 across the data by construction (x0 < min(x)); guard anyway so a
    # stray evaluation left of the pole yields the (huge but finite) wall value
    # instead of inf/nan.
    u = np.where(u <= 0.0, np.nan, u)
    ratio6 = (rm / u) ** 6
    return eps * (ratio6 * ratio6 - 2.0 * ratio6) + c


def _lj_force(x, eps, rm, x0, c):
    # F = -dV/dx = 12*eps*( rm^12 u^-13 - rm^6 u^-7 ), u = x - x0.
    u = np.asarray(x, dtype=float) - x0
    u = np.where(u <= 0.0, np.nan, u)
    ratio6 = (rm / u) ** 6
    return 12.0 * eps * (ratio6 * ratio6 - ratio6) / u


def _morse_exp(x, a, xe):
    # Clip the exponent so an optimiser probing large `a` can't overflow exp()
    # to inf (which would spam RuntimeWarnings and poison the least-squares
    # residual); t**2 stays finite for arg up to ~350.
    arg = np.clip(-a * (np.asarray(x, dtype=float) - xe), -350.0, 350.0)
    return np.exp(arg)


def _morse_value(x, de, a, xe, c):
    t = _morse_exp(x, a, xe)
    return de * (t * t - 2.0 * t) + c


def _morse_force(x, de, a, xe, c):
    # V = De(e^{-2a(x-xe)} - 2 e^{-a(x-xe)}) + c,  t = e^{-a(x-xe)}, dt/dx = -a t
    # dV/dx = -2 De a (t^2 - t)  ->  F = -dV/dx = 2 De a (t^2 - t)
    t = _morse_exp(x, a, xe)
    return 2.0 * de * a * (t * t - t)


# ---------------------------------------------------------------------------
# Fitting
# ---------------------------------------------------------------------------
def _clean_xy(x, y):
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    n = min(x.size, y.size)
    x, y = x[:n], y[:n]
    good = np.isfinite(x) & np.isfinite(y)
    x, y = x[good], y[good]
    order = np.argsort(x)
    return x[order], y[order]


def _baseline_and_well(x, y):
    """Far-field baseline c and (depth, position) of the well minimum.

    The far field is taken as the flat tail at large x (both potentials tend to
    a constant c there); the well is the global minimum.
    """
    n = y.size
    tail = max(3, n // 5)
    c0 = float(np.median(y[-tail:]))
    imin = int(np.argmin(y))
    depth = c0 - float(y[imin])
    return c0, depth, float(x[imin])


def _smooth(y):
    """Noise-suppressed copy of y, used only to locate features (never fitted)."""
    n = y.size
    win = max(5, (n // 25) | 1)
    if win >= n:
        win = n - 1 if (n - 1) % 2 else n - 2
    if win < 5:
        return y.copy()
    if _savgol is not None:
        try:
            return _savgol(y, win, 2, mode="interp")
        except Exception:
            pass
    pad = win // 2
    return np.convolve(np.pad(y, pad, mode="edge"), np.ones(win) / win, mode="valid")


def _noise_sigma(y, ys):
    r = y - ys
    mad = float(np.median(np.abs(r - np.median(r))))
    return 1.4826 * mad if mad > 0 else (float(np.std(r)) or 1e-12)


def robust_window(x, y, wall_frac=DEFAULT_WALL_FRAC):
    """First index to fit (x sorted ascending) in robust mode, plus info.

    Walks in from the far field (large x) on a smoothed copy and stops at the
    first attractive well followed by a real rise (more than noise and a
    quarter of the depth reached so far). From there the repulsive branch is
    kept only while it stays within ``wall_frac`` x depth of the minimum and
    keeps rising; the steep wall, the CO-bending turnover and any second dip
    beyond it are excluded. With no turning point (the curve never reaches the
    wall) everything is kept.
    """
    n = y.size
    ys = _smooth(y)
    sigma = _noise_sigma(y, ys)
    tail = max(3, n // 5)
    c0 = float(np.median(ys[-tail:]))
    run_min = np.inf
    i_min = n - 1
    i_well = None
    for i in range(n - 1, -1, -1):
        v = ys[i]
        if v < run_min:
            run_min, i_min = v, i
            continue
        depth = c0 - run_min
        if depth > 5.0 * sigma and v - run_min > max(5.0 * sigma, 0.25 * depth):
            i_well = i_min
            break
    info = {"sigma": sigma, "c0": c0, "well_index": i_well}
    if i_well is None:
        return 0, info
    y_well = float(ys[i_well])
    depth = c0 - y_well
    limit = y_well + max(float(wall_frac), 0.0) * depth
    run_max = y_well
    start = i_well
    for i in range(i_well - 1, -1, -1):
        v = ys[i]
        if v > limit:
            break
        if v < run_max - max(3.0 * sigma, 0.05 * depth):
            break  # the wall turned back down: CO bending / tip relaxation
        run_max = max(run_max, v)
        start = i
    return start, info


def fit_potential(x, y, model, robust=True, wall_frac=DEFAULT_WALL_FRAC):
    """Fit *model* ("Lennard-Jones" or "Morse") to (x, y).

    With *robust* (default) the repulsive branch is trimmed by
    :func:`robust_window`, residuals use a noise-scaled soft-L1 loss and the
    optimiser is multi-started (see the module docstring).

    Returns a result dict, or None if the fit is impossible (too few points,
    SciPy missing, or the optimiser failed). Never raises.

    Result keys: model, params (dict), param_order (tuple), func(x), force_func(x),
    rmse, x_min, x_max, singularity (float or None), trimmed (bool). x_min/x_max
    is the span actually fitted - in robust mode possibly narrower than the data.
    x_well/y_well/depth describe the fitted minimum; well_in_range is False when
    the curve never reached it (then they are clamped to x_min, i.e. bounds).
    """
    try:
        x, y = _clean_xy(x, y)
    except Exception:
        return None
    if x.size < 5 or _curve_fit is None:
        return None
    trimmed = False
    loss_kw = {}
    if robust:
        try:
            start, info = robust_window(x, y, wall_frac)
        except Exception:
            start, info = 0, {}
        if start > 0 and x.size - start >= 8:
            x, y = x[start:], y[start:]
            trimmed = True
        sigma = info.get("sigma")
        if sigma and np.isfinite(sigma):
            loss_kw = {"loss": "soft_l1", "f_scale": 2.0 * sigma}
    span = float(x[-1] - x[0])
    if not np.isfinite(span) or span <= 0:
        return None

    # Initial guesses from a smoothed copy: a raw argmin lands on a noise spike.
    c0, depth, x_at_min = _baseline_and_well(x, _smooth(y) if robust else y)
    eps0 = abs(depth) if depth != 0 else max(abs(c0), 1.0) * 0.1
    y_scale = max(eps0, np.ptp(y) or 1.0)
    # Pole/wall pinned just left of the data so the model is finite everywhere
    # it is ever drawn.
    x0_lo = x[0] - 5.0 * span
    x0_hi = x[0] - _POLE_MARGIN_FRAC * span

    if model == "Lennard-Jones":
        lo = [0.0, _POLE_MARGIN_FRAC * span, x0_lo, c0 - 10.0 * y_scale]
        hi = [1e6 * y_scale, 6.0 * span, x0_hi, c0 + 10.0 * y_scale]
        value_fn, force_fn = _lj_value, _lj_force
        order = ("eps", "rm", "x0", "c")
        starts = []
        for frac in ((0.03, 0.15, 0.5, 1.5) if robust else (0.15,)):
            x0_0 = x[0] - frac * span
            starts.append([eps0, max(x_at_min - x0_0, _POLE_MARGIN_FRAC * span), x0_0, c0])
    elif model == "Morse":
        lo = [0.0, 1e-6 / span, x[0] - span, c0 - 10.0 * y_scale]
        hi = [1e6 * y_scale, 1e6 / span, x[-1] + span, c0 + 10.0 * y_scale]
        value_fn, force_fn = _morse_value, _morse_force
        order = ("De", "a", "xe", "c")
        starts = [[eps0, k / span, x_at_min, c0] for k in ((2.0, 6.0, 20.0) if robust else (2.0,))]
    else:
        return None

    popt = None
    best = np.inf
    for p0 in starts:
        # Keep p0 strictly inside the bounds (curve_fit rejects p0 on a bound).
        p0 = [min(max(v, lo[i] + 1e-12), hi[i] - 1e-12) for i, v in enumerate(p0)]
        try:
            cand, _ = _curve_fit(value_fn, x, y, p0=p0, bounds=(lo, hi), maxfev=20000, **loss_kw)
            r = np.asarray(value_fn(x, *cand), dtype=float) - y
        except Exception:
            continue
        if loss_kw:
            z = (r / loss_kw["f_scale"]) ** 2
            cost = float(np.nansum(np.sqrt(1.0 + z) - 1.0))
        else:
            cost = float(np.nansum(r * r))
        if np.isfinite(cost) and cost < best:
            best, popt = cost, cand
    if popt is None:
        return None

    params = {name: float(val) for name, val in zip(order, popt)}
    try:
        resid = value_fn(x, *popt) - y
        rmse = float(np.sqrt(np.nanmean(resid ** 2)))
    except Exception:
        rmse = float("nan")
    if not np.all(np.isfinite(popt)):
        return None

    singularity = params.get("x0")  # only LJ has a real pole

    def func(xq, _p=popt, _f=value_fn):
        return _f(xq, *_p)

    def force_func(xq, _p=popt, _f=force_fn):
        return _f(xq, *_p)

    # Analytic well: LJ minimum at x0 + rm, Morse at xe, both of depth eps/De
    # below c. When it lies left of the fitted span the curve never reached
    # its minimum, so the reported well is clamped to the closest approach
    # (position and depth are then only bounds) and flagged.
    if model == "Lennard-Jones":
        x_well = params["x0"] + params["rm"]
        depth_fit = params["eps"]
    else:
        x_well = params["xe"]
        depth_fit = params["De"]
    well_in_range = bool(x[0] < x_well <= x[-1])
    if not well_in_range:
        x_well = float(min(max(x_well, x[0]), x[-1]))
        try:
            depth_fit = params["c"] - float(value_fn(np.array([x_well]), *popt)[0])
        except Exception:
            depth_fit = float("nan")
    y_well = params["c"] - depth_fit

    return {
        "model": model,
        "params": params,
        "param_order": order,
        "func": func,
        "force_func": force_func,
        "rmse": rmse,
        "x_min": float(x[0]),
        "x_max": float(x[-1]),
        "singularity": singularity,
        "trimmed": trimmed,
        "x_well": float(x_well),
        "y_well": float(y_well),
        "depth": float(depth_fit),
        "well_in_range": well_in_range,
    }


def params_summary(res):
    """Compact one-line parameter string for logs/legends."""
    if not res:
        return ""
    parts = [f"{k}={res['params'][k]:.4g}" for k in res.get("param_order", ())]
    rmse = res.get("rmse")
    if rmse is not None and np.isfinite(rmse):
        parts.append(f"RMSE={rmse:.3g}")
    return ", ".join(parts)
