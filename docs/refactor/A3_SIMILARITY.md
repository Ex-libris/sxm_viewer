# A3 - Near-duplicate method clusters

Methods >= 80% similar after normalization, grouped within each class/module. These are copy-paste families: each cluster is a candidate for one parameterized method, a shared helper, or a declarative table.

## Summary

- **111** clusters
- **249** methods involved
- **~2605** lines in clustered methods

| Methods | Lines | Owner | File | Members |
|---|---|---|---|---|
| 2 | 111 | CanvasImageItem | `sxm_viewer/gui/canvases/canvas_items.py` | _render_now, _render_vector_figure |
| 4 | 83 | (module) | `sxm_viewer/providers/nanonis/adapter.py` | _select_topo_axis, _select_z_axis, _select_bias_axis, _select_true_bias_axis |
| 4 | 75 | MatrixSpectroViewer | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py` | _build_slice_metric, _build_integral_metric, _build_peak_metric, _build_stat_metric |
| 2 | 64 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | _draw_shortcut_hint, _draw_acquisition_overlay |
| 7 | 63 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | set_show_shortcut_hint, set_show_profile_overlays, set_show_angle_overlays, set_show_title +3 |
| 2 | 61 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_review_low_conf_spectros, on_review_off_frame_spectros |
| 2 | 56 | (module) | `sxm_viewer/gui/viewer/loader.py` | _serialize_cache_value, _sanitize_metadata_value |
| 2 | 54 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_show_spectra_toggled, on_show_preview_spectra_toggled |
| 2 | 49 | (module) | `sxm_viewer/gui/plot_typography.py` | apply_qfont_style, apply_text_style |
| 2 | 47 | (module) | `sxm_viewer/gui/spectroscopy/popups.py` | _open_spectroscopy_compare_popup, _open_spectroscopy_popup |
| 2 | 46 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | load_files, load_folder |
| 2 | 40 | RecentFilesController | `sxm_viewer/gui/controllers/recent_files_controller.py` | _record_recent_session, _record_recent_dir |
| 3 | 39 | ExperimentalCanvasWindow | `sxm_viewer/gui/canvases/canvas_window.py` | _on_metadata_bar_toggled, _apply_global_show_colorbar, _apply_global_show_colorbar_ticks |
| 3 | 39 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_show_matrix_markers_toggled, on_show_single_markers_toggled, on_compact_markers_toggled |
| 2 | 38 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | _normalize_profile_marker_style, _normalize_profile_line_style |
| 2 | 37 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_show_spectro_miniatures_toggled, on_detail_grid_toggled |
| 2 | 36 | (module) | `sxm_viewer/config_io.py` | load_header_cache, load_collections_index |
| 3 | 36 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | _remember_closed_popup_profile_dialog, _remember_closed_main_profile_dialog, _remember_closed_canvas_window |
| 2 | 35 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | _dz_vs_previous_ch, _dz_vs_last_before_ch |
| 2 | 34 | ProfileDialog | `sxm_viewer/gui/dialogs/profile_dialog.py` | _deregister_workspace_dialog, _register_workspace_dialog |
| 3 | 33 | ExperimentalCanvasWindow | `sxm_viewer/gui/canvases/canvas_window.py` | _on_canvas_show_molecules_toggled, _on_metadata_unit_toggled, _on_global_show_title_toggled |
| 3 | 32 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_pick_spectro_stack_color, on_pick_spectro_matrix_color, on_pick_spectro_single_color |
| 2 | 32 | (module) | `sxm_viewer/gui/viewer/thumbnail_ui.py` | _make_thumb_release_handler, _make_spectro_thumb_release_handler |
| 2 | 30 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | _axis_coord_to_pixel_float, _axis_coord_to_pixel |
| 2 | 30 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | _prepare_profile_blit, _prepare_angle_blit |
| 2 | 29 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_add_selected_thumbnails_to_collection, on_add_selected_thumbnails_to_collection_picker |
| 2 | 28 | FilterController | `sxm_viewer/gui/controllers/filter_controller.py` | _canvas_filter_label, _canvas_filter_steps |
| 2 | 28 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | _load_molecule_overlay, _load_svg_molecule_overlay |
| 6 | 28 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | compare_menu_state, on_arrange_popouts, on_minimize_popouts, _update_matrix_summary_banner +2 |
| 2 | 28 | (module) | `sxm_viewer/gui/viewer/thumbnail_ui.py` | _refresh_spectro_thumb_selection_styles, _refresh_thumb_selection_styles |
| 2 | 27 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | _on_recent_molecules_updated, _on_recent_svg_molecules_updated |
| 2 | 27 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | _store_molecule_overlay, _store_svg_molecule_overlay |
| 2 | 27 | (module) | `sxm_viewer/gui/viewer/thumbnail_ui.py` | _make_thumb_move_handler, _make_spectro_thumb_move_handler |
| 2 | 26 | (module) | `sxm_viewer/config_io.py` | save_config, save_header_cache |
| 2 | 26 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_spectro_grid_as_matrix_toggled, on_spectro_force_single_toggled |
| 2 | 25 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | _blit_profile_artists, _blit_angle_frames |
| 2 | 25 | _CollectionQuickPickDialog | `sxm_viewer/gui/controllers/collection.py` | _prompt_new_collection, _prompt_browse_existing |
| 3 | 25 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_spec_coord_mode_changed, on_set_spectro_size, on_set_spectro_symbol |
| 3 | 24 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | _connect_scale_bar_events, _connect_angle_events, _connect_profile_events |
| 2 | 24 | PopupProfileController | `sxm_viewer/gui/controllers/profile.py` | _register_dialog, _deregister_dialog |

## Cluster detail

### CanvasImageItem - 2 similar methods (111 lines)

`sxm_viewer/gui/canvases/canvas_items.py`

- `_render_now` - line 334 (59 lines)
- `_render_vector_figure` - line 1330 (52 lines)

### (module-level) - 4 similar methods (83 lines)

`sxm_viewer/providers/nanonis/adapter.py`

- `_select_z_axis` - line 845 (23 lines)
- `_select_topo_axis` - line 879 (25 lines)
- `_select_bias_axis` - line 906 (14 lines)
- `_select_true_bias_axis` - line 922 (21 lines)

### MatrixSpectroViewer - 4 similar methods (75 lines)

`sxm_viewer/gui/dialogs/spectroscopy_dialogs.py`

- `_build_stat_metric` - line 5369 (16 lines)
- `_build_integral_metric` - line 5386 (18 lines)
- `_build_peak_metric` - line 5405 (17 lines)
- `_build_slice_metric` - line 5423 (24 lines)

### MultiPreviewCanvas - 2 similar methods (64 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_draw_acquisition_overlay` - line 9584 (30 lines)
- `_draw_shortcut_hint` - line 9723 (34 lines)

### MultiPreviewCanvas - 7 similar methods (63 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `set_show_title` - line 403 (9 lines)
- `set_show_acquisition_overlay` - line 413 (9 lines)
- `set_show_molecules` - line 423 (9 lines)
- `set_show_profile_overlays` - line 1147 (9 lines)
- `set_show_angle_overlays` - line 1157 (9 lines)
- `set_show_spectra_overlays` - line 1167 (8 lines)
- `set_show_shortcut_hint` - line 1176 (10 lines)

### SXMGridViewer - 2 similar methods (61 lines)

`sxm_viewer/gui/main_window.py`

- `on_review_low_conf_spectros` - line 8422 (32 lines)
- `on_review_off_frame_spectros` - line 8455 (29 lines)

### (module-level) - 2 similar methods (56 lines)

`sxm_viewer/gui/viewer/loader.py`

- `_serialize_cache_value` - line 309 (24 lines)
- `_sanitize_metadata_value` - line 348 (32 lines)

### SXMGridViewer - 2 similar methods (54 lines)

`sxm_viewer/gui/main_window.py`

- `on_show_spectra_toggled` - line 9738 (32 lines)
- `on_show_preview_spectra_toggled` - line 9795 (22 lines)

### (module-level) - 2 similar methods (49 lines)

`sxm_viewer/gui/plot_typography.py`

- `apply_text_style` - line 51 (24 lines)
- `apply_qfont_style` - line 77 (25 lines)

### (module-level) - 2 similar methods (47 lines)

`sxm_viewer/gui/spectroscopy/popups.py`

- `_open_spectroscopy_popup` - line 99 (20 lines)
- `_open_spectroscopy_compare_popup` - line 121 (27 lines)

### SXMGridViewer - 2 similar methods (46 lines)

`sxm_viewer/gui/main_window.py`

- `load_folder` - line 4656 (22 lines)
- `load_files` - line 4679 (24 lines)

### RecentFilesController - 2 similar methods (40 lines)

`sxm_viewer/gui/controllers/recent_files_controller.py`

- `_record_recent_dir` - line 42 (19 lines)
- `_record_recent_session` - line 124 (21 lines)

### ExperimentalCanvasWindow - 3 similar methods (39 lines)

`sxm_viewer/gui/canvases/canvas_window.py`

- `_on_metadata_bar_toggled` - line 730 (13 lines)
- `_apply_global_show_colorbar` - line 838 (13 lines)
- `_apply_global_show_colorbar_ticks` - line 852 (13 lines)

### SXMGridViewer - 3 similar methods (39 lines)

`sxm_viewer/gui/main_window.py`

- `on_show_matrix_markers_toggled` - line 9885 (13 lines)
- `on_show_single_markers_toggled` - line 9899 (13 lines)
- `on_compact_markers_toggled` - line 9913 (13 lines)

### MultiPreviewCanvas - 2 similar methods (38 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_normalize_profile_line_style` - line 3647 (17 lines)
- `_normalize_profile_marker_style` - line 3665 (21 lines)

### SXMGridViewer - 2 similar methods (37 lines)

`sxm_viewer/gui/main_window.py`

- `on_show_spectro_miniatures_toggled` - line 9771 (23 lines)
- `on_detail_grid_toggled` - line 9935 (14 lines)

### (module-level) - 2 similar methods (36 lines)

`sxm_viewer/config_io.py`

- `load_header_cache` - line 88 (22 lines)
- `load_collections_index` - line 157 (14 lines)

### SXMGridViewer - 3 similar methods (36 lines)

`sxm_viewer/gui/main_window.py`

- `_remember_closed_main_profile_dialog` - line 2784 (12 lines)
- `_remember_closed_popup_profile_dialog` - line 2797 (12 lines)
- `_remember_closed_canvas_window` - line 2844 (12 lines)

### SXMGridViewer - 2 similar methods (35 lines)

`sxm_viewer/gui/main_window.py`

- `_dz_vs_previous_ch` - line 6945 (17 lines)
- `_dz_vs_last_before_ch` - line 6963 (18 lines)

### ProfileDialog - 2 similar methods (34 lines)

`sxm_viewer/gui/dialogs/profile_dialog.py`

- `_register_workspace_dialog` - line 2008 (17 lines)
- `_deregister_workspace_dialog` - line 2026 (17 lines)

### ExperimentalCanvasWindow - 3 similar methods (33 lines)

`sxm_viewer/gui/canvases/canvas_window.py`

- `_on_canvas_show_molecules_toggled` - line 544 (11 lines)
- `_on_metadata_unit_toggled` - line 744 (11 lines)
- `_on_global_show_title_toggled` - line 756 (11 lines)

### SXMGridViewer - 3 similar methods (32 lines)

`sxm_viewer/gui/main_window.py`

- `on_pick_spectro_single_color` - line 9305 (10 lines)
- `on_pick_spectro_matrix_color` - line 9316 (11 lines)
- `on_pick_spectro_stack_color` - line 9328 (11 lines)

### (module-level) - 2 similar methods (32 lines)

`sxm_viewer/gui/viewer/thumbnail_ui.py`

- `_make_thumb_release_handler` - line 1458 (18 lines)
- `_make_spectro_thumb_release_handler` - line 1609 (14 lines)

### MultiPreviewCanvas - 2 similar methods (30 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_axis_coord_to_pixel` - line 10813 (15 lines)
- `_axis_coord_to_pixel_float` - line 10829 (15 lines)

### MultiPreviewCanvas - 2 similar methods (30 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_prepare_profile_blit` - line 6800 (15 lines)
- `_prepare_angle_blit` - line 6835 (15 lines)

### SXMGridViewer - 2 similar methods (29 lines)

`sxm_viewer/gui/main_window.py`

- `on_add_selected_thumbnails_to_collection` - line 6602 (14 lines)
- `on_add_selected_thumbnails_to_collection_picker` - line 6617 (15 lines)

### FilterController - 2 similar methods (28 lines)

`sxm_viewer/gui/controllers/filter_controller.py`

- `_canvas_filter_steps` - line 184 (13 lines)
- `_canvas_filter_label` - line 198 (15 lines)

### SXMGridViewer - 2 similar methods (28 lines)

`sxm_viewer/gui/main_window.py`

- `_load_molecule_overlay` - line 6424 (15 lines)
- `_load_svg_molecule_overlay` - line 6454 (13 lines)

### SXMGridViewer - 6 similar methods (28 lines)

`sxm_viewer/gui/main_window.py`

- `on_arrange_popouts` - line 6848 (5 lines)
- `on_minimize_popouts` - line 6854 (5 lines)
- `compare_menu_state` - line 6871 (6 lines)
- `_update_matrix_summary_banner` - line 8667 (4 lines)
- `_update_spec_selection_label` - line 9246 (4 lines)
- `_clear_multi_spec_selection` - line 9251 (4 lines)

### (module-level) - 2 similar methods (28 lines)

`sxm_viewer/gui/viewer/thumbnail_ui.py`

- `_refresh_spectro_thumb_selection_styles` - line 324 (14 lines)
- `_refresh_thumb_selection_styles` - line 1379 (14 lines)

