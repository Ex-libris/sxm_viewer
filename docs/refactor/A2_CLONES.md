# A2 - Structural clone detection

Code with identical *shape* after normalizing away identifiers, receivers and literals. Catches duplication that renaming hides. Method-call names are deliberately preserved, so a group here is usually directly collapsible into a shared helper.

## Whole-function clones (24 groups)

- **3x** `_spec_grid_row_col`, `MatrixSpectroViewer._spec_grid_row_col`, `spec_grid_row_col`
  - first: `sxm_viewer/gui/dialogs/matrix_fit.py:122` (29 lines)
- **3x** `_grid_local_pitch`, `MatrixSpectroViewer._grid_local_pitch`, `grid_local_pitch`
  - first: `sxm_viewer/gui/dialogs/matrix_fit.py:174` (24 lines)
- **2x** `extra_colormaps_status`, `extra_cmap_names`
  - first: `sxm_viewer/cmap_registry.py:199` (6 lines)
- **2x** `_coerce_value`, `_coerce_value`
  - first: `sxm_viewer/data/io.py:240` (10 lines)
- **2x** `_clean_channel_label`, `_sanitize_channel_label`
  - first: `sxm_viewer/data/spectroscopy.py:433` (6 lines)
- **2x** `MultiPreviewCanvas._pick_sb_text_color`, `MultiPreviewCanvas._pick_sb_bar_color`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:3574` (5 lines)
- **2x** `MultiPreviewCanvas._signature_key`, `SessionController._session_signature_key`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:9435` (9 lines)
- **2x** `MultiPreviewCanvas._crop_color_for_seq`, `QuickCropController._crop_color_for_seq`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:11799` (9 lines)
- **2x** `_CollectionTargetDialog._on_advanced_toggled`, `_CollectionQuickPickDialog._on_advanced_toggled`
  - first: `sxm_viewer/gui/controllers/collection.py:138` (3 lines)
- **2x** `_CollectionQuickPickDialog._display_label`, `CollectionBrowserDialog._display_label`
  - first: `sxm_viewer/gui/controllers/collection.py:262` (5 lines)
- **2x** `SingleFilterDialog._set_param_row_visible`, `CustomFilterDialog._set_param_row_visible`
  - first: `sxm_viewer/gui/dialogs/filters.py:363` (3 lines)
- **2x** `ProfileDialog._make_toggle_button`, `SpectroscopyPopup._make_toggle_button`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:1098` (14 lines)
- **2x** `ProfileDialog._set_advanced_options_visible`, `SpectroscopyPopup._set_advanced_options_visible`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:1113` (10 lines)
- **2x** `ProfileDialog.wheelEvent`, `SpectroscopyCompareDialog.wheelEvent`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:1182` (14 lines)
- **2x** `ProfileDialog.dragEnterEvent`, `ProfileDialog.dragMoveEvent`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:2207` (7 lines)
- **2x** `SpectroscopyPopup._populate_inset_settings_menu`, `SpectroscopyCompareDialog._populate_inset_settings_menu`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1377` (43 lines)
- **2x** `SpectroscopyPopup._pick_inset_marker_color`, `SpectroscopyCompareDialog._pick_inset_marker_color`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1443` (5 lines)
- **2x** `MatrixSpectroViewer.moveEvent`, `SpectroscopyCompareDialog.moveEvent`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:3976` (7 lines)
- **2x** `MatrixSpectroViewer._channel_unit_for_spec`, `SpectroscopyCompareDialog._channel_unit_for_spec`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4096` (9 lines)
- **2x** `MatrixSpectroViewer._grid_dims`, `grid_dims`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:5118` (16 lines)
- **2x** `SpectroscopyCompareDialog._get_icon`, `SpectroscopyCompareDialog._get_icon`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:6425` (4 lines)
- **2x** `_browser_payload_specs`, `SpectroSummaryDialog._payload_specs`
  - first: `sxm_viewer/gui/spectroscopy/browser.py:525` (14 lines)
- **2x** `_spec_identity_token`, `_spectro_identity_key`
  - first: `sxm_viewer/gui/viewer/loader.py:624` (16 lines)
- **2x** `_refresh_spectro_thumb_selection_styles`, `_refresh_thumb_selection_styles`
  - first: `sxm_viewer/gui/viewer/thumbnail_ui.py:324` (14 lines)

## Block clones (79 groups, >= 4 statements, >= 3 instances)

| Instances | Stmts | Total lines | Files | Example |
|---|---|---|---|---|
| 163 | 4 | 652 | 48 | `sxm_viewer/cmap_registry.py:368` |
| 46 | 5 | 230 | 24 | `sxm_viewer/config_io.py:36` |
| 28 | 4 | 112 | 21 | `sxm_viewer/data/spectroscopy.py:624` |
| 28 | 4 | 112 | 22 | `sxm_viewer/data/spectroscopy.py:482` |
| 24 | 4 | 96 | 13 | `sxm_viewer/config_defaults.py:1` |
| 22 | 4 | 88 | 16 | `sxm_viewer/data/spectroscopy.py:606` |
| 21 | 4 | 84 | 17 | `sxm_viewer/data/spectroscopy.py:615` |
| 12 | 6 | 72 | 6 | `sxm_viewer/data/spectroscopy.py:423` |
| 11 | 5 | 55 | 6 | `sxm_viewer/config_defaults.py:1` |
| 9 | 5 | 45 | 9 | `sxm_viewer/gui/canvases/canvas_items.py:209` |
| 8 | 5 | 40 | 8 | `sxm_viewer/data/spectroscopy.py:606` |
| 7 | 5 | 35 | 7 | `sxm_viewer/gui/canvases/canvas_rendering.py:515` |
| 5 | 7 | 35 | 3 | `sxm_viewer/gui/canvases/canvas_window_ui.py:95` |
| 8 | 4 | 32 | 8 | `sxm_viewer/config_io.py:117` |
| 8 | 4 | 32 | 7 | `sxm_viewer/gui/canvases/canvas_items.py:77` |
| 5 | 6 | 30 | 3 | `sxm_viewer/config_defaults.py:1` |
| 7 | 4 | 28 | 7 | `sxm_viewer/cmap_sorting.py:121` |
| 7 | 4 | 28 | 7 | `sxm_viewer/data/spectroscopy.py:768` |
| 7 | 4 | 28 | 6 | `sxm_viewer/gui/canvases/canvas_items.py:50` |
| 7 | 4 | 28 | 3 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:4013` |
| 7 | 4 | 28 | 6 | `sxm_viewer/gui/colormap_manager.py:45` |
| 6 | 4 | 24 | 6 | `sxm_viewer/gui/canvases/canvas_items.py:76` |
| 6 | 4 | 24 | 5 | `sxm_viewer/gui/canvases/canvas_items.py:314` |
| 6 | 4 | 24 | 6 | `sxm_viewer/gui/canvases/canvas_window_ui.py:209` |
| 4 | 5 | 20 | 4 | `sxm_viewer/data/spectroscopy.py:615` |
| 4 | 5 | 20 | 4 | `sxm_viewer/gui/canvases/canvas_rendering.py:191` |
| 5 | 4 | 20 | 4 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:5510` |
| 5 | 4 | 20 | 4 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:12011` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9744` |
| 4 | 5 | 20 | 3 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9745` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/molecular_overlay.py:487` |
| 5 | 4 | 20 | 4 | `sxm_viewer/gui/canvases/molecular_overlay.py:491` |
| 5 | 4 | 20 | 4 | `sxm_viewer/gui/canvases/svg_molecule_overlay.py:316` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/svg_molecule_overlay.py:317` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/controllers/profile.py:384` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/controllers/profile.py:462` |
| 3 | 6 | 18 | 3 | `sxm_viewer/gui/canvases/canvas_rendering.py:514` |
| 4 | 4 | 16 | 4 | `sxm_viewer/_shared.py:83` |
| 4 | 4 | 16 | 4 | `sxm_viewer/gui/canvases/canvas_items.py:394` |
| 4 | 4 | 16 | 4 | `sxm_viewer/gui/canvases/canvas_items.py:1122` |

<details><summary>163x 4-statement block (example sxm_viewer/cmap_registry.py:368)</summary>

- `sxm_viewer/cmap_registry.py:368`
- `sxm_viewer/config_io.py:36`
- `sxm_viewer/config_io.py:37`
- `sxm_viewer/data/spectroscopy.py:423`
- `sxm_viewer/data/spectroscopy.py:433`
- `sxm_viewer/data/spectroscopy.py:441`
- `sxm_viewer/data/spectroscopy.py:559`
- `sxm_viewer/data/spectroscopy.py:753`
- `sxm_viewer/gui/canvases/canvas_items.py:35`
- `sxm_viewer/gui/canvases/canvas_items.py:36`
- `sxm_viewer/gui/canvases/canvas_items.py:90`
- `sxm_viewer/gui/canvases/canvas_items.py:1005`
- `sxm_viewer/gui/canvases/canvas_items.py:1123`
- `sxm_viewer/gui/canvases/canvas_rendering.py:191`
- `sxm_viewer/gui/canvases/canvas_rendering.py:516`
- `sxm_viewer/gui/canvases/canvas_state.py:15`
- `sxm_viewer/gui/canvases/canvas_view.py:201`
- `sxm_viewer/gui/canvases/canvas_view.py:541`
- `sxm_viewer/gui/canvases/canvas_window.py:119`
- `sxm_viewer/gui/canvases/canvas_window.py:120`
- `sxm_viewer/gui/canvases/canvas_window.py:475`
- `sxm_viewer/gui/canvases/canvas_window.py:488`
- `sxm_viewer/gui/canvases/canvas_window_actions.py:7`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:79`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:95`

</details>

<details><summary>46x 5-statement block (example sxm_viewer/config_io.py:36)</summary>

- `sxm_viewer/config_io.py:36`
- `sxm_viewer/data/spectroscopy.py:423`
- `sxm_viewer/data/spectroscopy.py:433`
- `sxm_viewer/gui/canvases/canvas_items.py:35`
- `sxm_viewer/gui/canvases/canvas_window.py:119`
- `sxm_viewer/gui/canvases/canvas_window.py:475`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:95`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:97`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:98`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:158`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:558`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:218`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:302`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:303`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8939`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8940`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8941`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8942`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8943`
- `sxm_viewer/gui/colormap_gallery.py:417`
- `sxm_viewer/gui/colormap_gallery.py:432`
- `sxm_viewer/gui/controllers/collection.py:89`
- `sxm_viewer/gui/controllers/collection.py:494`
- `sxm_viewer/gui/controllers/preview_popup.py:163`
- `sxm_viewer/gui/controllers/quick_crop.py:553`

</details>

<details><summary>28x 4-statement block (example sxm_viewer/data/spectroscopy.py:624)</summary>

- `sxm_viewer/data/spectroscopy.py:88`
- `sxm_viewer/data/spectroscopy.py:473`
- `sxm_viewer/data/spectroscopy.py:624`
- `sxm_viewer/data/spectroscopy.py:718`
- `sxm_viewer/gui/canvases/canvas_items.py:209`
- `sxm_viewer/gui/canvases/canvas_rendering.py:330`
- `sxm_viewer/gui/canvases/canvas_rendering.py:514`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:258`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:956`
- `sxm_viewer/gui/colormap_gallery.py:349`
- `sxm_viewer/gui/dialogs/filters.py:586`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1451`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1512`
- `sxm_viewer/gui/dialogs/svg_molecule_style.py:157`
- `sxm_viewer/gui/main_window.py:936`
- `sxm_viewer/gui/main_window.py:5678`
- `sxm_viewer/gui/main_window.py:7678`
- `sxm_viewer/gui/main_window_layout.py:317`
- `sxm_viewer/gui/main_window_spectro.py:105`
- `sxm_viewer/gui/main_window_toolbar.py:330`
- `sxm_viewer/gui/minimap.py:77`
- `sxm_viewer/gui/thumbnail_render.py:235`
- `sxm_viewer/gui/viewer/export.py:278`
- `sxm_viewer/gui/viewer/measurement.py:163`
- `sxm_viewer/gui/viewer/thumbnail_ui.py:498`

</details>

<details><summary>28x 4-statement block (example sxm_viewer/data/spectroscopy.py:482)</summary>

- `sxm_viewer/data/spectroscopy.py:482`
- `sxm_viewer/gui/canvases/canvas_items.py:210`
- `sxm_viewer/gui/canvases/canvas_rendering.py:20`
- `sxm_viewer/gui/canvases/canvas_rendering.py:515`
- `sxm_viewer/gui/canvases/canvas_window.py:816`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:120`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:259`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:9745`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:10871`
- `sxm_viewer/gui/canvases/svg_molecule_overlay.py:1189`
- `sxm_viewer/gui/colormap_gallery.py:351`
- `sxm_viewer/gui/controllers/collection.py:361`
- `sxm_viewer/gui/controllers/session.py:168`
- `sxm_viewer/gui/controllers/session.py:1203`
- `sxm_viewer/gui/dialogs/filters.py:236`
- `sxm_viewer/gui/dialogs/profile_dialog.py:256`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1513`
- `sxm_viewer/gui/dialogs/svg_molecule_style.py:158`
- `sxm_viewer/gui/main_window.py:1256`
- `sxm_viewer/gui/main_window_layout.py:318`
- `sxm_viewer/gui/main_window_spectro.py:106`
- `sxm_viewer/gui/thumbnail_render.py:147`
- `sxm_viewer/gui/thumbnail_render.py:261`
- `sxm_viewer/processing/filters.py:137`
- `sxm_viewer/providers/nanonis/adapter.py:151`

</details>

<details><summary>24x 4-statement block (example sxm_viewer/config_defaults.py:1)</summary>

- `sxm_viewer/config_defaults.py:1`
- `sxm_viewer/config_defaults.py:2`
- `sxm_viewer/config_defaults.py:4`
- `sxm_viewer/config_defaults.py:6`
- `sxm_viewer/config_defaults.py:7`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:258`
- `sxm_viewer/gui/colormap_manager.py:36`
- `sxm_viewer/gui/constants.py:3`
- `sxm_viewer/gui/controllers/image_compare.py:212`
- `sxm_viewer/gui/controllers/session.py:792`
- `sxm_viewer/gui/dialogs/matrix_fit.py:424`
- `sxm_viewer/gui/dialogs/matrix_fit.py:425`
- `sxm_viewer/gui/main_window_toolbar.py:39`
- `sxm_viewer/gui/main_window_toolbar.py:40`
- `sxm_viewer/gui/spectroscopy/browser.py:652`
- `sxm_viewer/gui/spectroscopy/controller.py:1398`
- `sxm_viewer/gui/viewer/export.py:175`
- `sxm_viewer/gui/viewer/export.py:176`
- `sxm_viewer/processing/filters.py:348`
- `sxm_viewer/processing/filters.py:349`
- `sxm_viewer/processing/filters.py:353`
- `sxm_viewer/reporting/model.py:537`
- `sxm_viewer/reporting/model.py:539`
- `sxm_viewer/reporting/model.py:540`

</details>

<details><summary>22x 4-statement block (example sxm_viewer/data/spectroscopy.py:606)</summary>

- `sxm_viewer/data/spectroscopy.py:606`
- `sxm_viewer/gui/canvases/canvas_items.py:47`
- `sxm_viewer/gui/canvases/canvas_rendering.py:192`
- `sxm_viewer/gui/canvases/canvas_rendering.py:225`
- `sxm_viewer/gui/canvases/canvas_rendering.py:547`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:9746`
- `sxm_viewer/gui/compact_histogram.py:70`
- `sxm_viewer/gui/controllers/image_compare.py:1155`
- `sxm_viewer/gui/controllers/quick_crop.py:718`
- `sxm_viewer/gui/controllers/session.py:900`
- `sxm_viewer/gui/dialogs/svg_molecule_style.py:159`
- `sxm_viewer/gui/main_window.py:1649`
- `sxm_viewer/gui/main_window.py:4020`
- `sxm_viewer/gui/main_window.py:8313`
- `sxm_viewer/gui/main_window_layout.py:218`
- `sxm_viewer/gui/main_window_toolbar.py:24`
- `sxm_viewer/gui/spectroscopy/browser.py:218`
- `sxm_viewer/gui/spectroscopy/browser.py:361`
- `sxm_viewer/gui/thumbnail_render.py:197`
- `sxm_viewer/gui/thumbnail_render.py:274`
- `sxm_viewer/gui/viewer/loader.py:1109`
- `sxm_viewer/providers/nanonis/adapter.py:179`

</details>

<details><summary>21x 4-statement block (example sxm_viewer/data/spectroscopy.py:615)</summary>

- `sxm_viewer/data/spectroscopy.py:615`
- `sxm_viewer/gui/canvases/canvas_items.py:48`
- `sxm_viewer/gui/canvases/canvas_rendering.py:230`
- `sxm_viewer/gui/canvases/canvas_view.py:533`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:257`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:1771`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:9747`
- `sxm_viewer/gui/canvases/svg_molecule_overlay.py:835`
- `sxm_viewer/gui/controllers/image_compare.py:1156`
- `sxm_viewer/gui/controllers/session.py:1668`
- `sxm_viewer/gui/controllers/thumbnail_controller.py:2`
- `sxm_viewer/gui/dialogs/filters.py:254`
- `sxm_viewer/gui/main_window.py:7677`
- `sxm_viewer/gui/main_window.py:8314`
- `sxm_viewer/gui/main_window.py:8369`
- `sxm_viewer/gui/spectroscopy/controller.py:381`
- `sxm_viewer/gui/spectroscopy/summary_dialog.py:65`
- `sxm_viewer/gui/thumbnail_render.py:198`
- `sxm_viewer/gui/wsxm_stp.py:121`
- `sxm_viewer/processing/periodic_noise.py:102`
- `sxm_viewer/providers/nanonis/adapter.py:180`

</details>

<details><summary>12x 6-statement block (example sxm_viewer/data/spectroscopy.py:423)</summary>

- `sxm_viewer/data/spectroscopy.py:423`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:95`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:97`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:302`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8939`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8940`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8941`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8942`
- `sxm_viewer/gui/controllers/session.py:624`
- `sxm_viewer/gui/main_window_spectro.py:109`
- `sxm_viewer/gui/main_window_spectro.py:112`
- `sxm_viewer/gui/viewer/thumbnail_ui.py:532`

</details>

<details><summary>11x 5-statement block (example sxm_viewer/config_defaults.py:1)</summary>

- `sxm_viewer/config_defaults.py:1`
- `sxm_viewer/config_defaults.py:2`
- `sxm_viewer/config_defaults.py:4`
- `sxm_viewer/config_defaults.py:6`
- `sxm_viewer/gui/dialogs/matrix_fit.py:424`
- `sxm_viewer/gui/main_window_toolbar.py:39`
- `sxm_viewer/gui/viewer/export.py:175`
- `sxm_viewer/processing/filters.py:348`
- `sxm_viewer/processing/filters.py:349`
- `sxm_viewer/reporting/model.py:537`
- `sxm_viewer/reporting/model.py:539`

</details>

<details><summary>9x 5-statement block (example sxm_viewer/gui/canvases/canvas_items.py:209)</summary>

- `sxm_viewer/gui/canvases/canvas_items.py:209`
- `sxm_viewer/gui/canvases/canvas_rendering.py:514`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:258`
- `sxm_viewer/gui/colormap_gallery.py:349`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1512`
- `sxm_viewer/gui/dialogs/svg_molecule_style.py:157`
- `sxm_viewer/gui/main_window_layout.py:317`
- `sxm_viewer/gui/main_window_spectro.py:105`
- `sxm_viewer/gui/thumbnail_render.py:235`

</details>

<details><summary>8x 5-statement block (example sxm_viewer/data/spectroscopy.py:606)</summary>

- `sxm_viewer/data/spectroscopy.py:606`
- `sxm_viewer/gui/canvases/canvas_items.py:47`
- `sxm_viewer/gui/canvases/canvas_rendering.py:225`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:9746`
- `sxm_viewer/gui/controllers/image_compare.py:1155`
- `sxm_viewer/gui/main_window.py:8313`
- `sxm_viewer/gui/thumbnail_render.py:197`
- `sxm_viewer/providers/nanonis/adapter.py:179`

</details>

<details><summary>7x 5-statement block (example sxm_viewer/gui/canvases/canvas_rendering.py:515)</summary>

- `sxm_viewer/gui/canvases/canvas_rendering.py:515`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:120`
- `sxm_viewer/gui/colormap_gallery.py:351`
- `sxm_viewer/gui/controllers/collection.py:361`
- `sxm_viewer/gui/dialogs/filters.py:236`
- `sxm_viewer/gui/main_window_spectro.py:106`
- `sxm_viewer/reporting/model.py:288`

</details>

<details><summary>5x 7-statement block (example sxm_viewer/gui/canvases/canvas_window_ui.py:95)</summary>

- `sxm_viewer/gui/canvases/canvas_window_ui.py:95`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8939`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8940`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8941`
- `sxm_viewer/gui/main_window_spectro.py:109`

</details>

<details><summary>8x 4-statement block (example sxm_viewer/config_io.py:117)</summary>

- `sxm_viewer/config_io.py:117`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:557`
- `sxm_viewer/gui/controllers/spectro_compare.py:295`
- `sxm_viewer/gui/dialogs/image_adjust.py:267`
- `sxm_viewer/gui/main_window.py:753`
- `sxm_viewer/gui/minimap.py:180`
- `sxm_viewer/gui/viewer/thumbnail_ui.py:531`
- `sxm_viewer/reporting/channels.py:68`

</details>

<details><summary>8x 4-statement block (example sxm_viewer/gui/canvases/canvas_items.py:77)</summary>

- `sxm_viewer/gui/canvases/canvas_items.py:77`
- `sxm_viewer/gui/canvases/canvas_items.py:97`
- `sxm_viewer/gui/canvases/canvas_window.py:75`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:8411`
- `sxm_viewer/gui/controllers/session.py:537`
- `sxm_viewer/gui/main_window_layout.py:136`
- `sxm_viewer/gui/viewer/loader.py:1946`
- `sxm_viewer/providers/nanonis/adapter.py:1091`

</details>

