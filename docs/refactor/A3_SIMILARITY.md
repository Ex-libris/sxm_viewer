# A3 - Near-duplicate method clusters

Methods >= 80% similar after normalization, grouped within each class/module. These are copy-paste families: each cluster is a candidate for one parameterized method, a shared helper, or a declarative table.

## Summary

- **110** clusters
- **247** methods involved
- **~2627** lines in clustered methods

| Methods | Lines | Owner | File | Members |
|---|---|---|---|---|
| 2 | 111 | CanvasImageItem | `sxm_viewer/gui/canvases/canvas_items.py` | _render_now, _render_vector_figure |
| 4 | 83 | (module) | `sxm_viewer/providers/nanonis/adapter.py` | _select_topo_axis, _select_z_axis, _select_bias_axis, _select_true_bias_axis |
| 4 | 73 | MatrixSpectroViewer | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py` | _build_slice_metric, _build_peak_metric, _build_integral_metric, _build_stat_metric |
| 2 | 73 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_review_low_conf_spectros, on_review_off_frame_spectros |
| 2 | 64 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | _draw_shortcut_hint, _draw_acquisition_overlay |
| 7 | 63 | MultiPreviewCanvas | `sxm_viewer/gui/canvases/detail_preview_canvas.py` | set_show_shortcut_hint, set_show_profile_overlays, set_show_angle_overlays, set_show_title +3 |
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
| 2 | 37 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_unit_relative_toggled, on_preview_lock_toggled |
| 2 | 36 | (module) | `sxm_viewer/config_io.py` | load_header_cache, load_collections_index |
| 2 | 36 | SXMGridViewer | `sxm_viewer/gui/main_window.py` | on_scale_bar_toggled, on_unit_display_toggled |
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

### MatrixSpectroViewer - 4 similar methods (73 lines)

`sxm_viewer/gui/dialogs/spectroscopy_dialogs.py`

- `_build_stat_metric` - line 5367 (16 lines)
- `_build_integral_metric` - line 5384 (16 lines)
- `_build_peak_metric` - line 5401 (17 lines)
- `_build_slice_metric` - line 5419 (24 lines)

### SXMGridViewer - 2 similar methods (73 lines)

`sxm_viewer/gui/main_window.py`

- `on_review_low_conf_spectros` - line 9460 (38 lines)
- `on_review_off_frame_spectros` - line 9499 (35 lines)

### MultiPreviewCanvas - 2 similar methods (64 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_draw_acquisition_overlay` - line 9802 (30 lines)
- `_draw_shortcut_hint` - line 9941 (34 lines)

### MultiPreviewCanvas - 7 similar methods (63 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `set_show_title` - line 402 (9 lines)
- `set_show_acquisition_overlay` - line 412 (9 lines)
- `set_show_molecules` - line 422 (9 lines)
- `set_show_profile_overlays` - line 1146 (9 lines)
- `set_show_angle_overlays` - line 1156 (9 lines)
- `set_show_spectra_overlays` - line 1166 (8 lines)
- `set_show_shortcut_hint` - line 1175 (10 lines)

### (module-level) - 2 similar methods (56 lines)

`sxm_viewer/gui/viewer/loader.py`

- `_serialize_cache_value` - line 309 (24 lines)
- `_sanitize_metadata_value` - line 348 (32 lines)

### SXMGridViewer - 2 similar methods (54 lines)

`sxm_viewer/gui/main_window.py`

- `on_show_spectra_toggled` - line 11836 (32 lines)
- `on_show_preview_spectra_toggled` - line 11893 (22 lines)

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

- `load_folder` - line 5164 (22 lines)
- `load_files` - line 5187 (24 lines)

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

- `on_show_matrix_markers_toggled` - line 11983 (13 lines)
- `on_show_single_markers_toggled` - line 11997 (13 lines)
- `on_compact_markers_toggled` - line 12011 (13 lines)

### MultiPreviewCanvas - 2 similar methods (38 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_normalize_profile_line_style` - line 3646 (17 lines)
- `_normalize_profile_marker_style` - line 3664 (21 lines)

### SXMGridViewer - 2 similar methods (37 lines)

`sxm_viewer/gui/main_window.py`

- `on_show_spectro_miniatures_toggled` - line 11869 (23 lines)
- `on_detail_grid_toggled` - line 12033 (14 lines)

### SXMGridViewer - 2 similar methods (37 lines)

`sxm_viewer/gui/main_window.py`

- `on_unit_relative_toggled` - line 6475 (19 lines)
- `on_preview_lock_toggled` - line 11565 (18 lines)

### (module-level) - 2 similar methods (36 lines)

`sxm_viewer/config_io.py`

- `load_header_cache` - line 88 (22 lines)
- `load_collections_index` - line 157 (14 lines)

### SXMGridViewer - 2 similar methods (36 lines)

`sxm_viewer/gui/main_window.py`

- `on_unit_display_toggled` - line 6456 (18 lines)
- `on_scale_bar_toggled` - line 6561 (18 lines)

### SXMGridViewer - 3 similar methods (36 lines)

`sxm_viewer/gui/main_window.py`

- `_remember_closed_main_profile_dialog` - line 3279 (12 lines)
- `_remember_closed_popup_profile_dialog` - line 3292 (12 lines)
- `_remember_closed_canvas_window` - line 3339 (12 lines)

### SXMGridViewer - 2 similar methods (35 lines)

`sxm_viewer/gui/main_window.py`

- `_dz_vs_previous_ch` - line 7654 (17 lines)
- `_dz_vs_last_before_ch` - line 7672 (18 lines)

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

- `on_pick_spectro_single_color` - line 11391 (10 lines)
- `on_pick_spectro_matrix_color` - line 11402 (11 lines)
- `on_pick_spectro_stack_color` - line 11414 (11 lines)

### (module-level) - 2 similar methods (32 lines)

`sxm_viewer/gui/viewer/thumbnail_ui.py`

- `_make_thumb_release_handler` - line 1458 (18 lines)
- `_make_spectro_thumb_release_handler` - line 1609 (14 lines)

### MultiPreviewCanvas - 2 similar methods (30 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_axis_coord_to_pixel` - line 11086 (15 lines)
- `_axis_coord_to_pixel_float` - line 11102 (15 lines)

### MultiPreviewCanvas - 2 similar methods (30 lines)

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- `_prepare_profile_blit` - line 6799 (15 lines)
- `_prepare_angle_blit` - line 6834 (15 lines)

### SXMGridViewer - 2 similar methods (29 lines)

`sxm_viewer/gui/main_window.py`

- `on_add_selected_thumbnails_to_collection` - line 7247 (14 lines)
- `on_add_selected_thumbnails_to_collection_picker` - line 7262 (15 lines)

### FilterController - 2 similar methods (28 lines)

`sxm_viewer/gui/controllers/filter_controller.py`

- `_canvas_filter_steps` - line 184 (13 lines)
- `_canvas_filter_label` - line 198 (15 lines)

### SXMGridViewer - 2 similar methods (28 lines)

`sxm_viewer/gui/main_window.py`

- `_load_molecule_overlay` - line 7066 (15 lines)
- `_load_svg_molecule_overlay` - line 7096 (13 lines)

