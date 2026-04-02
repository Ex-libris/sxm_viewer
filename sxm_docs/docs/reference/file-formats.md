# Supported File Formats

SXM Viewer is built around SPM imaging and spectroscopy data from Anfatec/Omicron workflows, with additional support for MATRIX and Nanonis-related formats.

---

## Imaging formats

| Format | Notes |
|---|---|
| Anfatec SXM | Primary imaging workflow with multi-channel support |
| MATRIX | Vendor import supported |
| Nanonis | Adapter-based support has been added for scans and related metadata |
| WSxM XYZ | Import and export support |

---

## Spectroscopy formats

Supported spectroscopy workflows include:

- single-point spectroscopy traces
- matrix / grid spectroscopy
- KPFM-related spectroscopy content
- WSxM XYZ export paths

The parser history also shows hardening for Omicron `.dat` spectroscopy files and clearer rejection of malformed inputs.

---

## Notes

!!! note
    The exact set of imported channels can depend on the provider and the metadata available in the source files.

!!! note
    Some support paths are implemented through adapters that normalize external vendor formats into the viewer's main workflow.

---

## Related pages

- [Loading Data](../browsing/loading.md)
- [Spectroscopy Overview](../spectroscopy/overview.md)
- [Matrix Scans](../spectroscopy/matrix.md)
