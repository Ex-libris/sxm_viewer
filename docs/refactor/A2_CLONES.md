# A2 - Structural clone detection

Code with identical *shape* after normalizing away identifiers, receivers and literals. Catches duplication that renaming hides. Method-call names are deliberately preserved, so a group here is usually directly collapsible into a shared helper.

## Whole-function clones (31 groups)

- **5x** `TestComposeAxisLabel.test_strips_bracketed_unit_already_in_name`, `TestComposeAxisLabel.test_strips_parenthesized_unit_already_in_name`, `TestComposeAxisLabel.test_leaves_unrelated_brackets_alone`, `TestComposeAxisLabel.test_blank_unit_leaves_name_untouched`, `TestSiFormat.test_blank_name`
  - first: `sxm_viewer/tests/test_profile_units.py:88` (5 lines)
- **3x** `_spec_grid_row_col`, `MatrixSpectroViewer._spec_grid_row_col`, `spec_grid_row_col`
  - first: `sxm_viewer/gui/dialogs/matrix_fit.py:29` (29 lines)
- **3x** `_grid_local_pitch`, `MatrixSpectroViewer._grid_local_pitch`, `grid_local_pitch`
  - first: `sxm_viewer/gui/dialogs/matrix_fit.py:81` (24 lines)
- **3x** `_unique_path`, `_unique`, `_unique`
  - first: `sxm_viewer/gui/viewer/export.py:103` (8 lines)
- **3x** `TestSiScale.test_pico_range_current`, `TestSiScale.test_nano_range_distance`, `TestSiScale.test_kilo_range`
  - first: `sxm_viewer/tests/test_profile_units.py:122` (4 lines)
- **2x** `extra_colormaps_status`, `extra_cmap_names`
  - first: `sxm_viewer/cmap_registry.py:236` (6 lines)
- **2x** `_coerce_value`, `_coerce_value`
  - first: `sxm_viewer/data/io.py:240` (10 lines)
- **2x** `_clean_channel_label`, `_sanitize_channel_label`
  - first: `sxm_viewer/data/spectroscopy.py:423` (6 lines)
- **2x** `_finite`, `_finite_or_none`
  - first: `sxm_viewer/figures/height_series.py:169` (6 lines)
- **2x** `MultiPreviewCanvas._pick_sb_text_color`, `MultiPreviewCanvas._pick_sb_bar_color`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:4265` (5 lines)
- **2x** `MultiPreviewCanvas._signature_key`, `SessionController._session_signature_key`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:10509` (9 lines)
- **2x** `MultiPreviewCanvas._crop_color_for_seq`, `QuickCropController._crop_color_for_seq`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:12622` (9 lines)
- **2x** `_CollectionTargetDialog._on_advanced_toggled`, `_CollectionQuickPickDialog._on_advanced_toggled`
  - first: `sxm_viewer/gui/controllers/collection.py:138` (3 lines)
- **2x** `_CollectionQuickPickDialog._display_label`, `CollectionBrowserDialog._display_label`
  - first: `sxm_viewer/gui/controllers/collection.py:262` (5 lines)
- **2x** `SingleFilterDialog._set_param_row_visible`, `CustomFilterDialog._set_param_row_visible`
  - first: `sxm_viewer/gui/dialogs/filters.py:261` (3 lines)
- **2x** `_grid_local_orientation`, `MatrixSpectroViewer._grid_local_orientation`
  - first: `sxm_viewer/gui/dialogs/matrix_fit.py:107` (38 lines)
- **2x** `ProfileDialog.wheelEvent`, `SpectroscopyCompareDialog.wheelEvent`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:1310` (14 lines)
- **2x** `ProfileDialog.dragEnterEvent`, `ProfileDialog.dragMoveEvent`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:3339` (7 lines)
- **2x** `_get_profile_figure_preset`, `get_figure_layout_preset`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:108` (6 lines)
- **2x** `SpectroscopyPopup._populate_inset_settings_menu`, `SpectroscopyCompareDialog._populate_inset_settings_menu`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1905` (3 lines)
- **2x** `SpectroscopyPopup._populate_inset_cmap_submenu`, `SpectroscopyCompareDialog._populate_inset_cmap_submenu`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1909` (18 lines)
- **2x** `SpectroscopyPopup._pick_inset_marker_color`, `SpectroscopyCompareDialog._pick_inset_marker_color`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1968` (5 lines)
- **2x** `MatrixSpectroViewer.moveEvent`, `SpectroscopyCompareDialog.moveEvent`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4494` (7 lines)
- **2x** `MatrixSpectroViewer._channel_unit_for_spec`, `SpectroscopyCompareDialog._channel_unit_for_spec`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4614` (9 lines)
- **2x** `MatrixSpectroViewer._grid_dims`, `grid_dims`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:5636` (16 lines)
- **2x** `MenuPrefs.set_pinned`, `MenuPrefs.record`
  - first: `sxm_viewer/gui/menu_kit.py:126` (4 lines)
- **2x** `_browser_payload_specs`, `SpectroSummaryDialog._payload_specs`
  - first: `sxm_viewer/gui/spectroscopy/browser.py:504` (14 lines)
- **2x** `_spec_identity_token`, `_spectro_identity_key`
  - first: `sxm_viewer/gui/viewer/loader.py:567` (16 lines)
- **2x** `_refresh_spectro_thumb_selection_styles`, `_refresh_thumb_selection_styles`
  - first: `sxm_viewer/gui/viewer/thumbnail_ui.py:303` (14 lines)
- **2x** `_cache_dir_for`, `_cache_dir_for`
  - first: `sxm_viewer/providers/nanonis/adapter.py:411` (7 lines)
- **2x** `TestFmtTol.test_never_claims_more_precision_than_a_bare_format`, `TestFmtTol.test_small_current_value_gets_pico_prefix`
  - first: `sxm_viewer/tests/test_profile_measure.py:29` (7 lines)

## Block clones (74 groups, >= 4 statements, >= 3 instances)

| Instances | Stmts | Total lines | Files | Example |
|---|---|---|---|---|
| 10 | 4 | 40 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901` |
| 8 | 5 | 40 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901` |
| 9 | 4 | 36 | 6 | `sxm_viewer/gui/canvases/canvas_window_ui.py:115` |
| 6 | 6 | 36 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901` |
| 4 | 8 | 32 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333` |
| 5 | 6 | 30 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325` |
| 5 | 6 | 30 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 7 | 4 | 28 | 7 | `sxm_viewer/data/spectroscopy.py:182` |
| 4 | 7 | 28 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901` |
| 7 | 4 | 28 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325` |
| 7 | 4 | 28 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 4 | 7 | 28 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333` |
| 4 | 7 | 28 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10334` |
| 5 | 5 | 25 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325` |
| 5 | 5 | 25 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10326` |
| 5 | 5 | 25 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 5 | 5 | 25 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 5 | 5 | 25 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10345` |
| 6 | 4 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333` |
| 6 | 4 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 4 | 6 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333` |
| 4 | 6 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10334` |
| 4 | 6 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 3 | 8 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10334` |
| 3 | 8 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336` |
| 3 | 8 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336` |
| 3 | 8 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336` |
| 3 | 8 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 3 | 8 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 3 | 8 | 24 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10345` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344` |
| 3 | 7 | 21 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10345` |
| 5 | 4 | 20 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10326` |
| 5 | 4 | 20 | 1 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10326` |

<details><summary>10x 4-statement block (example sxm_viewer/gui/canvases/detail_preview_canvas.py:12901)</summary>

- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12902`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12903`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12904`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12905`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13024`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13025`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13026`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13027`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13028`

</details>

<details><summary>8x 5-statement block (example sxm_viewer/gui/canvases/detail_preview_canvas.py:12901)</summary>

- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12902`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12903`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12904`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13024`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13025`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13026`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13027`

</details>

<details><summary>9x 4-statement block (example sxm_viewer/gui/canvases/canvas_window_ui.py:115)</summary>

- `sxm_viewer/gui/canvases/canvas_window_ui.py:115`
- `sxm_viewer/gui/dialogs/profile_dialog.py:365`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1624`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:6639`
- `sxm_viewer/gui/main_window.py:317`
- `sxm_viewer/gui/main_window.py:318`
- `sxm_viewer/gui/main_window.py:3909`
- `sxm_viewer/gui/viewer/thumbnail_ui.py:572`
- `sxm_viewer/tests/test_profile_units.py:45`

</details>

<details><summary>6x 6-statement block (example sxm_viewer/gui/canvases/detail_preview_canvas.py:12901)</summary>

- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12902`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12903`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13024`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13025`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13026`

</details>

<details><summary>4x 8-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10345`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10364`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10378`

</details>

<details><summary>5x 6-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10348`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10367`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10381`

</details>

<details><summary>5x 6-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10364`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10377`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10448`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10476`

</details>

<details><summary>7x 4-statement block (example sxm_viewer/data/spectroscopy.py:182)</summary>

- `sxm_viewer/data/spectroscopy.py:182`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:9667`
- `sxm_viewer/gui/minimap.py:243`
- `sxm_viewer/gui/thumbnail_render.py:560`
- `sxm_viewer/gui/viewer/thumbnail_ui.py:552`
- `sxm_viewer/gui/virtual_copies.py:45`
- `sxm_viewer/tests/test_profile_units.py:36`

</details>

<details><summary>4x 7-statement block (example sxm_viewer/gui/canvases/detail_preview_canvas.py:12901)</summary>

- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12901`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:12902`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13024`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:13025`

</details>

<details><summary>7x 4-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10348`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10367`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10381`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10761`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10788`

</details>

<details><summary>7x 4-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10344`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10377`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10447`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10516`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10798`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10845`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10846`

</details>

<details><summary>4x 7-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10333`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10345`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10364`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10378`

</details>

<details><summary>4x 7-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10334)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10334`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10346`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10365`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10379`

</details>

<details><summary>5x 5-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10325`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10348`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10367`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10381`

</details>

<details><summary>5x 5-statement block (example sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10326)</summary>

- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10326`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10336`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10348`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10367`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10381`

</details>

