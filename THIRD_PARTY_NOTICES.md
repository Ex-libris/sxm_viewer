# Third-party code, formats and references

SXM Viewer's own code is released under the MIT License (see [LICENSE](LICENSE)).
This file lists material in the repository that is under other terms, and the
work SXM Viewer builds on.

## Bundled code under its own license

### nanonispy2 1.2.0 (MIT License)

`sxm_viewer/providers/nanonis/vendor/nanonispy2-1.2.0/` is an unmodified copy
of [nanonispy2](https://pypi.org/project/nanonispy2/) by Julian Ceddia
(<https://github.com/ceds92/nanonispy2>), itself a fork of
[nanonispy](https://github.com/underchemist/nanonispy) by Yann-Sebastien
Tremblay-Johnston. Its license is kept in that folder.

### Force-deconvolution MATLAB scripts (CC BY 2.0)

`scripts/matrixF.m` and `scripts/saderF.m` are Copyright 2011 Joachim Welker,
Esther Illek and Franz J. Giessibl, distributed under the
[Creative Commons Attribution 2.0 License](https://creativecommons.org/licenses/by/2.0/).
They accompany:

> J. Welker, E. Illek, F. J. Giessibl, "Analysis of force-deconvolution methods
> in frequency-modulation atomic force microscopy",
> *Beilstein J. Nanotechnol.* **3**, 238–248 (2012).
> <https://doi.org/10.3762/bjnano.3.27>

They are included unmodified, for reference; the application does not call
them.

## Methods implemented

The "Convert to force" tool in the spectroscopy comparison dialog is an
independent Python implementation of:

- J. E. Sader, S. P. Jarvis, "Accurate formulas for interaction force and
  energy in frequency modulation force spectroscopy",
  *Appl. Phys. Lett.* **84**, 1801 (2004). <https://doi.org/10.1063/1.1667267>
- F. J. Giessibl, "A direct method to calculate tip–sample forces from
  frequency shifts in frequency-modulation atomic force microscopy",
  *Appl. Phys. Lett.* **78**, 123 (2001). <https://doi.org/10.1063/1.1335546>

## File formats

- **Anfatec / Omicron** header and channel files: parsed following the file
  structure defined by Anfatec Instruments AG (Oelsnitz, Germany,
  <http://www.anfatec.de/>), as documented in their
  [Python examples for SXM](https://www.anfatec.net/downloads/SXM/Python_Examples_for_SXM.zip).
  SXM Viewer's parser is an independent implementation; no Anfatec code is
  included.
- **Nanonis** `.sxm`, `.dat`, `.3ds`: read with nanonispy2 (above).
- **WSxM** `.stp` / XYZ export: written in the WSxM file format so results can
  be opened in WSxM (<https://wsxm.eu/>). If you use WSxM, please cite
  I. Horcas et al., *Rev. Sci. Instrum.* **78**, 013705 (2007),
  <https://doi.org/10.1063/1.2432410>.

## Inspiration

- L. M. A. Perdigão, "LMAPper – Where Scanning Probe Microscopy and Molecular
  Visualisation Meet" (pre-print),
  <https://sourceforge.net/p/spm-and-mol-viewer/wiki/Pre-print/>.
- WSxM (<https://wsxm.eu/>).
- A. Riss, "SpmImage Tycoon: Organize and analyze scanning probe microscopy
  data", *J. Open Source Softw.* 7(77), 4644 (2022),
  <https://doi.org/10.21105/joss.04644>;
  <https://github.com/alexriss/SpmImageTycoon.jl>.

No code from any of these projects is included (LMAPper and SpmImage Tycoon
are not written in Python, and WSxM is closed source).

## Runtime dependencies

These are installed separately and are not distributed with SXM Viewer: NumPy,
SciPy, Matplotlib (BSD/PSF-style), scikit-image (BSD), OpenCV (Apache 2.0),
pywin32 (PSF), and optionally imageio (BSD), Pillow (MIT-CMU) and
[colormaps](https://pratiman-91.github.io/colormaps/) by Pratiman Patel (MIT).

**PyQt5** is licensed under the GPL v3 (or a commercial license from Riverbank
Computing). Distributing SXM Viewer as source under MIT is fine, but any
bundled binary that includes PyQt5 (for example a PyInstaller build) has to
comply with the GPL.
