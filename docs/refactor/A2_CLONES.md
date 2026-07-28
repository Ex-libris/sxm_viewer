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
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:3575` (5 lines)
- **2x** `MultiPreviewCanvas._signature_key`, `SessionController._session_signature_key`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:9439` (9 lines)
- **2x** `MultiPreviewCanvas._crop_color_for_seq`, `QuickCropController._crop_color_for_seq`
  - first: `sxm_viewer/gui/canvases/detail_preview_canvas.py:11526` (9 lines)
- **2x** `_CollectionTargetDialog._on_advanced_toggled`, `_CollectionQuickPickDialog._on_advanced_toggled`
  - first: `sxm_viewer/gui/controllers/collection.py:138` (3 lines)
- **2x** `_CollectionQuickPickDialog._display_label`, `CollectionBrowserDialog._display_label`
  - first: `sxm_viewer/gui/controllers/collection.py:262` (5 lines)
- **2x** `SingleFilterDialog._set_param_row_visible`, `CustomFilterDialog._set_param_row_visible`
  - first: `sxm_viewer/gui/dialogs/filters.py:363` (3 lines)
- **2x** `_grid_local_orientation`, `MatrixSpectroViewer._grid_local_orientation`
  - first: `sxm_viewer/gui/dialogs/matrix_fit.py:200` (38 lines)
- **2x** `ProfileDialog._make_toggle_button`, `SpectroscopyPopup._make_toggle_button`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:1098` (14 lines)
- **2x** `ProfileDialog._set_advanced_options_visible`, `SpectroscopyPopup._set_advanced_options_visible`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:1113` (10 lines)
- **2x** `ProfileDialog.wheelEvent`, `SpectroscopyCompareDialog.wheelEvent`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:1182` (14 lines)
- **2x** `ProfileDialog.dragEnterEvent`, `ProfileDialog.dragMoveEvent`
  - first: `sxm_viewer/gui/dialogs/profile_dialog.py:2207` (7 lines)
- **2x** `SpectroscopyPopup._populate_inset_settings_menu`, `SpectroscopyCompareDialog._populate_inset_settings_menu`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1379` (43 lines)
- **2x** `SpectroscopyPopup._pick_inset_marker_color`, `SpectroscopyCompareDialog._pick_inset_marker_color`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1445` (5 lines)
- **2x** `MatrixSpectroViewer.moveEvent`, `SpectroscopyCompareDialog.moveEvent`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:3978` (7 lines)
- **2x** `MatrixSpectroViewer._channel_unit_for_spec`, `SpectroscopyCompareDialog._channel_unit_for_spec`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4098` (9 lines)
- **2x** `MatrixSpectroViewer._grid_dims`, `grid_dims`
  - first: `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:5120` (16 lines)
- **2x** `_browser_payload_specs`, `SpectroSummaryDialog._payload_specs`
  - first: `sxm_viewer/gui/spectroscopy/browser.py:525` (14 lines)
- **2x** `_spec_identity_token`, `_spectro_identity_key`
  - first: `sxm_viewer/gui/viewer/loader.py:624` (16 lines)
- **2x** `_refresh_spectro_thumb_selection_styles`, `_refresh_thumb_selection_styles`
  - first: `sxm_viewer/gui/viewer/thumbnail_ui.py:324` (14 lines)

## Block clones (91 groups, >= 4 statements, >= 3 instances)

| Instances | Stmts | Total lines | Files | Example |
|---|---|---|---|---|
| 144 | 4 | 576 | 45 | `sxm_viewer/_shared.py:23` |
| 39 | 5 | 195 | 17 | `sxm_viewer/data/io.py:1` |
| 32 | 4 | 128 | 23 | `sxm_viewer/cmap_sorting.py:179` |
| 31 | 4 | 124 | 23 | `sxm_viewer/cmap_sorting.py:176` |
| 25 | 4 | 100 | 23 | `sxm_viewer/cmap_sorting.py:177` |
| 24 | 4 | 96 | 19 | `sxm_viewer/cmap_sorting.py:178` |
| 11 | 6 | 66 | 9 | `sxm_viewer/data/io.py:1` |
| 12 | 5 | 60 | 10 | `sxm_viewer/cmap_sorting.py:176` |
| 12 | 5 | 60 | 9 | `sxm_viewer/gui/canvases/canvas_items.py:959` |
| 12 | 4 | 48 | 10 | `sxm_viewer/gui/colormap_gallery.py:429` |
| 9 | 4 | 36 | 7 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:1441` |
| 9 | 4 | 36 | 8 | `sxm_viewer/gui/main_window.py:1128` |
| 8 | 4 | 32 | 7 | `sxm_viewer/cmap_sorting.py:175` |
| 8 | 4 | 32 | 8 | `sxm_viewer/data/spectroscopy.py:527` |
| 8 | 4 | 32 | 7 | `sxm_viewer/gui/canvases/canvas_window_ui.py:101` |
| 8 | 4 | 32 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:3154` |
| 8 | 4 | 32 | 7 | `sxm_viewer/gui/dialogs/image_adjust.py:611` |
| 6 | 5 | 30 | 5 | `sxm_viewer/cmap_sorting.py:178` |
| 5 | 6 | 30 | 4 | `sxm_viewer/gui/canvases/canvas_window.py:86` |
| 6 | 5 | 30 | 6 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:331` |
| 6 | 5 | 30 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:3154` |
| 7 | 4 | 28 | 5 | `sxm_viewer/cmap_registry.py:124` |
| 5 | 5 | 25 | 5 | `sxm_viewer/cmap_sorting.py:177` |
| 5 | 5 | 25 | 5 | `sxm_viewer/gui/canvases/canvas_window_ui.py:99` |
| 5 | 5 | 25 | 5 | `sxm_viewer/gui/main_window.py:1127` |
| 6 | 4 | 24 | 6 | `sxm_viewer/gui/canvases/canvas_items.py:300` |
| 6 | 4 | 24 | 6 | `sxm_viewer/gui/canvases/canvas_view.py:498` |
| 4 | 6 | 24 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:3154` |
| 6 | 4 | 24 | 6 | `sxm_viewer/gui/controllers/session.py:1154` |
| 5 | 4 | 20 | 4 | `sxm_viewer/data/spectroscopy.py:378` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/canvas_items.py:299` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/canvas_items.py:652` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/canvas_items.py:653` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/canvas_window.py:810` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:1079` |
| 5 | 4 | 20 | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8909` |
| 5 | 4 | 20 | 5 | `sxm_viewer/gui/canvases/preview_export_figures.py:300` |
| 5 | 4 | 20 | 4 | `sxm_viewer/gui/colormap_manager.py:49` |
| 5 | 4 | 20 | 4 | `sxm_viewer/gui/controllers/recent_files_controller.py:8` |
| 3 | 6 | 18 | 3 | `sxm_viewer/cmap_sorting.py:177` |

<details><summary>144x 4-statement block (example sxm_viewer/_shared.py:23)</summary>

- `sxm_viewer/_shared.py:23`
- `sxm_viewer/app_meta.py:13`
- `sxm_viewer/data/io.py:1`
- `sxm_viewer/data/io.py:17`
- `sxm_viewer/data/io.py:19`
- `sxm_viewer/data/spectroscopy.py:348`
- `sxm_viewer/data/spectroscopy.py:629`
- `sxm_viewer/gui/canvases/canvas_items.py:46`
- `sxm_viewer/gui/canvases/canvas_items.py:52`
- `sxm_viewer/gui/canvases/canvas_items.py:662`
- `sxm_viewer/gui/canvases/canvas_items.py:959`
- `sxm_viewer/gui/canvases/canvas_window.py:86`
- `sxm_viewer/gui/canvases/canvas_window.py:143`
- `sxm_viewer/gui/canvases/canvas_window.py:144`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:71`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:72`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:73`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:98`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:147`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:148`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:149`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:176`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:177`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:178`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:216`

</details>

<details><summary>39x 5-statement block (example sxm_viewer/data/io.py:1)</summary>

- `sxm_viewer/data/io.py:1`
- `sxm_viewer/data/io.py:17`
- `sxm_viewer/gui/canvases/canvas_window.py:143`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:71`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:72`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:147`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:148`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:176`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:177`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:261`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:332`
- `sxm_viewer/gui/controllers/collection.py:320`
- `sxm_viewer/gui/controllers/histogram.py:61`
- `sxm_viewer/gui/controllers/quick_crop.py:79`
- `sxm_viewer/gui/controllers/quick_crop.py:134`
- `sxm_viewer/gui/dialogs/image_adjust.py:279`
- `sxm_viewer/gui/dialogs/image_adjust.py:580`
- `sxm_viewer/gui/dialogs/image_adjust.py:581`
- `sxm_viewer/gui/dialogs/profile_dialog.py:612`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1483`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1495`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:851`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:852`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7481`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7488`

</details>

<details><summary>32x 4-statement block (example sxm_viewer/cmap_sorting.py:179)</summary>

- `sxm_viewer/cmap_sorting.py:179`
- `sxm_viewer/gui/canvases/canvas_rendering.py:193`
- `sxm_viewer/gui/canvases/canvas_window.py:297`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:269`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:331`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:11485`
- `sxm_viewer/gui/controllers/image_compare.py:291`
- `sxm_viewer/gui/controllers/recent_files_controller.py:100`
- `sxm_viewer/gui/controllers/report.py:447`
- `sxm_viewer/gui/controllers/session.py:782`
- `sxm_viewer/gui/dialogs/filters.py:145`
- `sxm_viewer/gui/dialogs/image_adjust.py:768`
- `sxm_viewer/gui/dialogs/matrix_fit.py:287`
- `sxm_viewer/gui/dialogs/profile_dialog.py:225`
- `sxm_viewer/gui/dialogs/profile_dialog.py:498`
- `sxm_viewer/gui/dialogs/profile_dialog.py:986`
- `sxm_viewer/gui/dialogs/profile_dialog.py:2354`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:3563`
- `sxm_viewer/gui/figure_layout_presets.py:58`
- `sxm_viewer/gui/figure_layout_presets.py:72`
- `sxm_viewer/gui/profile_links.py:45`
- `sxm_viewer/gui/profile_links.py:106`
- `sxm_viewer/gui/spectroscopy/browser.py:312`
- `sxm_viewer/gui/spectroscopy/browser.py:321`
- `sxm_viewer/gui/spectroscopy/controller.py:1489`

</details>

<details><summary>31x 4-statement block (example sxm_viewer/cmap_sorting.py:176)</summary>

- `sxm_viewer/cmap_sorting.py:176`
- `sxm_viewer/gui/canvases/canvas_items.py:960`
- `sxm_viewer/gui/canvases/canvas_window.py:87`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:99`
- `sxm_viewer/gui/canvases/preview_export_figures.py:78`
- `sxm_viewer/gui/colormap_gallery.py:427`
- `sxm_viewer/gui/controllers/collection.py:93`
- `sxm_viewer/gui/controllers/collection.py:328`
- `sxm_viewer/gui/controllers/collection.py:1227`
- `sxm_viewer/gui/controllers/quick_crop.py:300`
- `sxm_viewer/gui/controllers/quick_crop.py:477`
- `sxm_viewer/gui/controllers/spectro_compare.py:386`
- `sxm_viewer/gui/dialogs/filters.py:674`
- `sxm_viewer/gui/dialogs/image_adjust.py:566`
- `sxm_viewer/gui/dialogs/image_adjust.py:583`
- `sxm_viewer/gui/dialogs/profile_dialog.py:500`
- `sxm_viewer/gui/dialogs/profile_dialog.py:614`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1515`
- `sxm_viewer/gui/main_window.py:133`
- `sxm_viewer/gui/main_window.py:812`
- `sxm_viewer/gui/main_window.py:1124`
- `sxm_viewer/gui/main_window_spectro.py:201`
- `sxm_viewer/gui/main_window_state.py:162`
- `sxm_viewer/gui/minimap.py:208`
- `sxm_viewer/gui/profile_links.py:51`

</details>

<details><summary>25x 4-statement block (example sxm_viewer/cmap_sorting.py:177)</summary>

- `sxm_viewer/cmap_sorting.py:177`
- `sxm_viewer/gui/canvases/canvas_window.py:88`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:289`
- `sxm_viewer/gui/canvases/preview_axes_sync.py:28`
- `sxm_viewer/gui/canvases/preview_export_figures.py:79`
- `sxm_viewer/gui/canvases/svg_molecule_overlay.py:172`
- `sxm_viewer/gui/colormap_gallery.py:320`
- `sxm_viewer/gui/controllers/collection.py:94`
- `sxm_viewer/gui/controllers/collection.py:329`
- `sxm_viewer/gui/controllers/histogram.py:72`
- `sxm_viewer/gui/controllers/image_compare.py:102`
- `sxm_viewer/gui/controllers/quick_crop.py:301`
- `sxm_viewer/gui/dialogs/filters.py:686`
- `sxm_viewer/gui/dialogs/matrix_fit.py:261`
- `sxm_viewer/gui/dialogs/profile_dialog.py:501`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1517`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:362`
- `sxm_viewer/gui/main_window_spectro.py:202`
- `sxm_viewer/gui/main_window_toolbar.py:285`
- `sxm_viewer/gui/spectroscopy/browser.py:310`
- `sxm_viewer/gui/spectroscopy/controller.py:109`
- `sxm_viewer/gui/spectroscopy/popups.py:166`
- `sxm_viewer/gui/viewer/measurement.py:331`
- `sxm_viewer/gui/viewer/preview.py:148`
- `sxm_viewer/utils/units.py:42`

</details>

<details><summary>24x 4-statement block (example sxm_viewer/cmap_sorting.py:178)</summary>

- `sxm_viewer/cmap_sorting.py:178`
- `sxm_viewer/gui/canvases/canvas_layout.py:10`
- `sxm_viewer/gui/canvases/canvas_window.py:296`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:158`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:872`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:5427`
- `sxm_viewer/gui/colormap_gallery.py:401`
- `sxm_viewer/gui/controllers/collection.py:735`
- `sxm_viewer/gui/controllers/histogram.py:73`
- `sxm_viewer/gui/dialogs/filters.py:293`
- `sxm_viewer/gui/dialogs/matrix_fit.py:262`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1257`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:3562`
- `sxm_viewer/gui/drift_animation.py:313`
- `sxm_viewer/gui/main_window.py:217`
- `sxm_viewer/gui/main_window.py:717`
- `sxm_viewer/gui/main_window.py:10332`
- `sxm_viewer/gui/main_window_layout.py:262`
- `sxm_viewer/gui/plot_typography.py:131`
- `sxm_viewer/gui/spectroscopy/browser.py:311`
- `sxm_viewer/gui/spectroscopy/browser.py:320`
- `sxm_viewer/gui/spectroscopy/controller.py:578`
- `sxm_viewer/gui/spectroscopy/popups.py:167`
- `sxm_viewer/reporting/pdf.py:505`

</details>

<details><summary>11x 6-statement block (example sxm_viewer/data/io.py:1)</summary>

- `sxm_viewer/data/io.py:1`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:71`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:147`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:176`
- `sxm_viewer/gui/dialogs/image_adjust.py:580`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1483`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:851`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7481`
- `sxm_viewer/gui/drift_animation.py:247`
- `sxm_viewer/gui/main_window_toolbar.py:247`
- `sxm_viewer/gui/spectroscopy/browser.py:792`

</details>

<details><summary>12x 5-statement block (example sxm_viewer/cmap_sorting.py:176)</summary>

- `sxm_viewer/cmap_sorting.py:176`
- `sxm_viewer/gui/canvases/canvas_window.py:87`
- `sxm_viewer/gui/canvases/preview_export_figures.py:78`
- `sxm_viewer/gui/controllers/collection.py:93`
- `sxm_viewer/gui/controllers/collection.py:328`
- `sxm_viewer/gui/controllers/quick_crop.py:300`
- `sxm_viewer/gui/dialogs/filters.py:674`
- `sxm_viewer/gui/dialogs/profile_dialog.py:500`
- `sxm_viewer/gui/dialogs/profile_dialog.py:1515`
- `sxm_viewer/gui/main_window_spectro.py:201`
- `sxm_viewer/gui/spectroscopy/controller.py:107`
- `sxm_viewer/utils/units.py:41`

</details>

<details><summary>12x 5-statement block (example sxm_viewer/gui/canvases/canvas_items.py:959)</summary>

- `sxm_viewer/gui/canvases/canvas_items.py:959`
- `sxm_viewer/gui/canvases/canvas_window.py:86`
- `sxm_viewer/gui/canvases/canvas_window_ui.py:98`
- `sxm_viewer/gui/controllers/collection.py:91`
- `sxm_viewer/gui/controllers/collection.py:327`
- `sxm_viewer/gui/controllers/spectro_compare.py:377`
- `sxm_viewer/gui/dialogs/image_adjust.py:582`
- `sxm_viewer/gui/dialogs/profile_dialog.py:499`
- `sxm_viewer/gui/dialogs/profile_dialog.py:613`
- `sxm_viewer/gui/main_window.py:132`
- `sxm_viewer/gui/main_window.py:811`
- `sxm_viewer/utils/units.py:39`

</details>

<details><summary>12x 4-statement block (example sxm_viewer/gui/colormap_gallery.py:429)</summary>

- `sxm_viewer/gui/colormap_gallery.py:429`
- `sxm_viewer/gui/dialogs/matrix_fit.py:652`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:6098`
- `sxm_viewer/gui/main_window.py:1127`
- `sxm_viewer/gui/main_window.py:1194`
- `sxm_viewer/gui/main_window_state.py:378`
- `sxm_viewer/gui/main_window_toolbar.py:24`
- `sxm_viewer/gui/main_window_toolbar.py:131`
- `sxm_viewer/gui/theme.py:653`
- `sxm_viewer/gui/viewer/loader.py:961`
- `sxm_viewer/gui/viewer/preview.py:271`
- `sxm_viewer/providers/nanonis/adapter.py:324`

</details>

<details><summary>9x 4-statement block (example sxm_viewer/gui/canvases/detail_preview_canvas.py:1441)</summary>

- `sxm_viewer/gui/canvases/detail_preview_canvas.py:1441`
- `sxm_viewer/gui/canvases/detail_preview_canvas.py:5030`
- `sxm_viewer/gui/canvases/preview_export_figures.py:66`
- `sxm_viewer/gui/controllers/report.py:1`
- `sxm_viewer/gui/dialogs/profile_dialog.py:334`
- `sxm_viewer/gui/main_window.py:407`
- `sxm_viewer/gui/main_window.py:8217`
- `sxm_viewer/gui/thumbnail_render.py:135`
- `sxm_viewer/reporting/model.py:72`

</details>

<details><summary>9x 4-statement block (example sxm_viewer/gui/main_window.py:1128)</summary>

- `sxm_viewer/gui/main_window.py:1128`
- `sxm_viewer/gui/main_window_toolbar.py:25`
- `sxm_viewer/gui/spectroscopy/overlays.py:87`
- `sxm_viewer/gui/spectroscopy/summary_dialog.py:232`
- `sxm_viewer/gui/theme.py:676`
- `sxm_viewer/gui/viewer/preview.py:554`
- `sxm_viewer/gui/wsxm_stp.py:266`
- `sxm_viewer/providers/nanonis/adapter.py:176`
- `sxm_viewer/providers/nanonis/adapter.py:325`

</details>

<details><summary>8x 4-statement block (example sxm_viewer/cmap_sorting.py:175)</summary>

- `sxm_viewer/cmap_sorting.py:175`
- `sxm_viewer/data/__init__.py:1`
- `sxm_viewer/gui/canvases/preview_export_figures.py:77`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:3593`
- `sxm_viewer/gui/main_window.py:408`
- `sxm_viewer/gui/main_window.py:800`
- `sxm_viewer/gui/minimap.py:207`
- `sxm_viewer/gui/viewer/thumbnail_ui.py:663`

</details>

<details><summary>8x 4-statement block (example sxm_viewer/data/spectroscopy.py:527)</summary>

- `sxm_viewer/data/spectroscopy.py:527`
- `sxm_viewer/gui/canvases/canvas_rendering.py:375`
- `sxm_viewer/gui/colormap_gallery.py:229`
- `sxm_viewer/gui/dialogs/matrix_fit.py:884`
- `sxm_viewer/gui/dialogs/profile_dialog.py:645`
- `sxm_viewer/gui/main_window.py:405`
- `sxm_viewer/gui/main_window_layout.py:528`
- `sxm_viewer/processing/filters.py:14`

</details>

<details><summary>8x 4-statement block (example sxm_viewer/gui/canvases/canvas_window_ui.py:101)</summary>

- `sxm_viewer/gui/canvases/canvas_window_ui.py:101`
- `sxm_viewer/gui/colormap_gallery.py:428`
- `sxm_viewer/gui/dialogs/image_adjust.py:584`
- `sxm_viewer/gui/dialogs/profile_dialog.py:886`
- `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:425`
- `sxm_viewer/gui/main_window.py:1126`
- `sxm_viewer/gui/viewer/export.py:333`
- `sxm_viewer/gui/viewer/export.py:518`

</details>

