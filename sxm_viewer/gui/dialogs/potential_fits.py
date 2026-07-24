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
"""
from __future__ import annotations

import numpy as np

try:  # SciPy is an optional dependency elsewhere in the app; degrade cleanly.
    from scipy.optimize import curve_fit as _curve_fit
except Exception:  # pragma: no cover - exercised only where SciPy is absent
    _curve_fit = None


MODELS = ("Lennard-Jones", "Morse")

# Smallest fraction of the data span the pole is allowed to approach the data.
_POLE_MARGIN_FRAC = 1e-3


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


def fit_potential(x, y, model):
    """Fit *model* ("Lennard-Jones" or "Morse") to (x, y).

    Returns a result dict, or None if the fit is impossible (too few points,
    SciPy missing, or the optimiser failed). Never raises.

    Result keys: model, params (dict), param_order (tuple), func(x), force_func(x),
    rmse, x_min, x_max, singularity (float or None).
    """
    try:
        x, y = _clean_xy(x, y)
    except Exception:
        return None
    if x.size < 5 or _curve_fit is None:
        return None
    span = float(x[-1] - x[0])
    if not np.isfinite(span) or span <= 0:
        return None

    c0, depth, x_at_min = _baseline_and_well(x, y)
    eps0 = abs(depth) if depth != 0 else max(abs(c0), 1.0) * 0.1
    y_scale = max(eps0, np.ptp(y) or 1.0)
    # Pole/wall pinned just left of the data so the model is finite everywhere
    # it is ever drawn.
    x0_lo = x[0] - 5.0 * span
    x0_hi = x[0] - _POLE_MARGIN_FRAC * span

    try:
        if model == "Lennard-Jones":
            x0_0 = x[0] - 0.15 * span
            rm0 = max(x_at_min - x0_0, _POLE_MARGIN_FRAC * span)
            p0 = [eps0, rm0, x0_0, c0]
            lo = [0.0, _POLE_MARGIN_FRAC * span, x0_lo, c0 - 10.0 * y_scale]
            hi = [1e6 * y_scale, 6.0 * span, x0_hi, c0 + 10.0 * y_scale]
            value_fn, force_fn = _lj_value, _lj_force
            order = ("eps", "rm", "x0", "c")
        elif model == "Morse":
            p0 = [eps0, 2.0 / span, x_at_min, c0]
            lo = [0.0, 1e-6 / span, x[0] - span, c0 - 10.0 * y_scale]
            hi = [1e6 * y_scale, 1e6 / span, x[-1] + span, c0 + 10.0 * y_scale]
            value_fn, force_fn = _morse_value, _morse_force
            order = ("De", "a", "xe", "c")
        else:
            return None

        # Keep p0 strictly inside the bounds (curve_fit rejects p0 on a bound).
        p0 = [min(max(v, lo[i] + 1e-12), hi[i] - 1e-12) for i, v in enumerate(p0)]
        popt, _ = _curve_fit(value_fn, x, y, p0=p0, bounds=(lo, hi), maxfev=20000)
    except Exception:
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
