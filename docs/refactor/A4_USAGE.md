# A4 - Usage and reachability

Name-based call analysis. **Never called** is a review queue, not a delete list - Qt signal connections and `getattr` dispatch are invisible here (flagged where a matching string literal exists).

## True phantom calls (never defined, never assigned)

None found.

## Injected callbacks (51 names)

> Called as `self.x(...)` with no `def x`, but assigned somewhere as an attribute - a deliberate plug-in point. Not bugs, but an **implicit, untyped interface**: nothing declares them, so a typo or a rename fails silently at runtime. Worth documenting as a real extension surface.

| Callback | Call sites | First site |
|---|---|---|
| `_compare_menu_callback` | 7 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9146` |
| `_value_callback` | 4 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:10694` |
| `_label_scale_cb` | 4 | `sxm_viewer/gui/dialogs/profile_dialog.py:491` |
| `_profile_marker_callback` | 3 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:3941` |
| `_profile_highlight_cb` | 3 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:5875` |
| `_preview_callback` | 3 | `sxm_viewer/gui/dialogs/filters.py:466` |
| `_copy_feedback_handler` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:4228` |
| `_profile_state_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7043` |
| `_views_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:724` |
| `_crop_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:10811` |
| `_last_svg_molecule_dir_cb` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7336` |
| `_histogram_auto_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:4456` |
| `_apply_popup_style_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9087` |
| `_collection_menu_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9131` |
| `apply_step` | 2 | `sxm_viewer/gui/dialogs/filters.py:453` |
| `_marker_update_cb` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:2595` |
| `_marker_key_cb` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:2465` |
| `_highlight_overlay_cb` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:2756` |
| `SpectroSummaryDialog` | 2 | `sxm_viewer/gui/main_window_spectro.py:159` |
| `_usage_provider` | 1 | `sxm_viewer/cmap_sorting.py:301` |
| `angle_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:4209` |
| `_font_change_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:1140` |
| `profile_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7029` |
| `_svg_molecule_style_defaults_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7540` |
| `_molecule_palette_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7939` |
| `_filter_menu_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8788` |
| `_fixed_crop_history_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:11629` |
| `_double_click_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7110` |
| `_spectra_click_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7158` |
| `_recent_svg_molecule_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7492` |

## Never referenced (72 definitions)

- 58 with no matching string literal (strongest delete candidates)
- 14 whose name appears as a string somewhere (likely dynamic use - verify before touching)

| Name | Location | Lines | Owner |
|---|---|---|---|
| `_spawn_lazy_popup_shell` | `sxm_viewer/gui/controllers/session.py:2168` | 94 | SessionController |
| `_build_spec_transform` | `sxm_viewer/gui/viewer/preview.py:72` | 30 | (module) |
| `_fast_spec_load_data` | `sxm_viewer/providers/nanonis/adapter.py:339` | 29 | (module) |
| `_safe_resize_event` | `sxm_viewer/_shared.py:43` | 21 | (module) |
| `toggle_measure` | `sxm_viewer/gui/controllers/profile.py:68` | 20 | PopupProfileController |
| `grouped_base_cmap_names` | `sxm_viewer/cmap_registry.py:359` | 18 | (module) |
| `on_show_crop_history_overlay_toggled` | `sxm_viewer/gui/main_window.py:12273` | 17 | SXMGridViewer |
| `grouped_cmap_names` | `sxm_viewer/cmap_registry.py:242` | 16 | (module) |
| `focus_first_matrix_dataset` | `sxm_viewer/gui/controllers/thumbnail_controller.py:141` | 15 | ThumbnailController |
| `reveal_points_for_file` | `sxm_viewer/gui/main_window.py:2705` | 15 | SXMGridViewer |
| `_current_visible_entries` | `sxm_viewer/gui/spectroscopy/summary_dialog.py:339` | 14 | SpectroSummaryDialog |
| `_alignment_score` | `sxm_viewer/gui/controllers/image_compare.py:319` | 12 | (module) |
| `_current_selected_spec` | `sxm_viewer/gui/spectroscopy/summary_dialog.py:327` | 11 | SpectroSummaryDialog |
| `bond_pairs` | `sxm_viewer/gui/canvases/svg_molecule_overlay.py:1186` | 10 | SvgMoleculeOverlay |
| `close_tracked_popups` | `sxm_viewer/gui/controllers/quick_crop.py:774` | 10 | QuickCropController |
| `on_spec_coord_mode_changed` | `sxm_viewer/gui/main_window.py:11374` | 9 | SXMGridViewer |
| `_site_sort_tuple` | `sxm_viewer/gui/spectroscopy/controller.py:642` | 9 | (module) |
| `_channel_labels` | `sxm_viewer/data/spectroscopy.py:423` | 8 | (module) |
| `_scale_bar_width` | `sxm_viewer/gui/canvases/canvas_items.py:320` | 8 | CanvasImageItem |
| `_run_pending_spectro_load` | `sxm_viewer/gui/main_window.py:9062` | 8 | SXMGridViewer |
| `stage` | `sxm_viewer/utils/logging.py:17` | 8 | (module) |
| `_normalized_value` | `sxm_viewer/gui/canvases/canvas_rendering.py:161` | 7 | (module) |
| `_create_icon_button` | `sxm_viewer/gui/canvases/canvas_window.py:161` | 7 | ExperimentalCanvasWindow |
| `_text_scale_factor` | `sxm_viewer/gui/canvases/canvas_items.py:898` | 6 | CanvasImageItem |
| `on_spec_invert_changed` | `sxm_viewer/gui/main_window.py:11384` | 6 | SXMGridViewer |
| `unregister_profile_canvas` | `sxm_viewer/gui/profile_links.py:31` | 6 | (module) |
| `hoverEnterEvent` | `sxm_viewer/gui/canvases/canvas_items.py:496` | 5 | CanvasImageItem |
| `_hline` | `sxm_viewer/gui/canvases/canvas_window.py:212` | 5 | ExperimentalCanvasWindow |
| `_add_menu_widget` | `sxm_viewer/gui/main_window_layout.py:20` | 5 | (module) |
| `hide_entry` | `sxm_viewer/gui/minimap.py:41` | 5 | FrameMiniMap |
| `_open_single` | `sxm_viewer/gui/spectroscopy/summary_dialog.py:261` | 5 | SpectroSummaryDialog |
| `dump_metadata_json` | `sxm_viewer/cmap_sorting.py:242` | 4 | (module) |
| `_on_preview_toggle` | `sxm_viewer/gui/dialogs/profile_dialog.py:1814` | 4 | ProfileDialog |
| `_xyz_filename` | `sxm_viewer/gui/main_window.py:8005` | 4 | SXMGridViewer |
| `_update_spec_selection_label` | `sxm_viewer/gui/main_window.py:10430` | 4 | SXMGridViewer |
| `extra_cmap_names` | `sxm_viewer/cmap_registry.py:207` | 3 | (module) |
| `set_grid_size` | `sxm_viewer/gui/canvases/canvas_view.py:37` | 3 | CanvasGraphicsView |
| `set_measurement_shortcuts_enabled` | `sxm_viewer/gui/canvases/detail_preview_canvas.py:503` | 3 | MultiPreviewCanvas |
| `clear_views` | `sxm_viewer/gui/canvases/detail_preview_canvas.py:982` | 3 | MultiPreviewCanvas |
| `set_reversed` | `sxm_viewer/gui/colormap_manager.py:68` | 3 | ColormapManager |
| `pending_cmap_name` | `sxm_viewer/gui/colormap_manager.py:112` | 3 | ColormapManager |
| `pending_colormap` | `sxm_viewer/gui/colormap_manager.py:119` | 3 | ColormapManager |
| `_apply_dark_mode` | `sxm_viewer/gui/main_window.py:2008` | 3 | SXMGridViewer |
| `on_dark_mode_toggled` | `sxm_viewer/gui/main_window.py:4778` | 3 | SXMGridViewer |
| `log_progress` | `sxm_viewer/utils/logging.py:11` | 3 | (module) |
| `hoverLeaveEvent` | `sxm_viewer/gui/canvases/canvas_items.py:502` | 2 | CanvasImageItem |
| `_create_toolbar_section` | `sxm_viewer/gui/canvases/canvas_window.py:169` | 2 | ExperimentalCanvasWindow |
| `_create_toolbar_group` | `sxm_viewer/gui/canvases/canvas_window.py:175` | 2 | ExperimentalCanvasWindow |
| `_create_separator` | `sxm_viewer/gui/canvases/canvas_window.py:178` | 2 | ExperimentalCanvasWindow |
| `get_cpk_color` | `sxm_viewer/gui/canvases/molecular_overlay.py:180` | 2 | (module) |
| `set_view` | `sxm_viewer/gui/colormap_gallery.py:158` | 2 | _CmapCardDelegate |
| `strategy` | `sxm_viewer/gui/colormap_gallery.py:370` | 2 | ColormapGallery |
| `applied_name` | `sxm_viewer/gui/colormap_manager.py:91` | 2 | ColormapManager |
| `applied_reversed` | `sxm_viewer/gui/colormap_manager.py:95` | 2 | ColormapManager |
| `applied_cmap_name` | `sxm_viewer/gui/colormap_manager.py:116` | 2 | ColormapManager |
| `restore_popups` | `sxm_viewer/gui/controllers/quick_crop.py:771` | 2 | QuickCropController |
| `get_result_maps` | `sxm_viewer/gui/dialogs/matrix_fit.py:711` | 2 | MatrixFitDialog |
| `on_save_session_as` | `sxm_viewer/gui/main_window.py:3684` | 2 | SXMGridViewer |

## Called exactly once (503 definitions)

> Inline candidates - strongest for short shims.

82 of these are <= 4 lines.

| Name | Defined | Lines | Only caller |
|---|---|---|---|
| `project_root` | `sxm_viewer/app_meta.py:27` | 2 | `sxm_viewer/app_meta.py:32` |
| `samples_dir` | `sxm_viewer/app_meta.py:31` | 2 | `sxm_viewer/app_meta.py:36` |
| `_luminance` | `sxm_viewer/cmap_sorting.py:93` | 3 | `sxm_viewer/cmap_sorting.py:138` |
| `_split_tokens` | `sxm_viewer/data/spectroscopy.py:553` | 3 | `sxm_viewer/data/spectroscopy.py:95` |
| `data_array` | `sxm_viewer/gui/canvases/canvas_items.py:1427` | 2 | `sxm_viewer/gui/canvases/canvas_window.py:426` |
| `safe_default_filename` | `sxm_viewer/gui/canvases/canvas_rendering.py:74` | 4 | `sxm_viewer/gui/canvases/canvas_items.py:1404` |
| `canvas_items` | `sxm_viewer/gui/canvases/canvas_state.py:11` | 2 | `sxm_viewer/gui/canvases/canvas_state.py:20` |
| `selected_canvas_items` | `sxm_viewer/gui/canvases/canvas_state.py:15` | 2 | `sxm_viewer/gui/canvases/canvas_state.py:63` |
| `_on_export_image` | `sxm_viewer/gui/canvases/canvas_window.py:1503` | 2 | `sxm_viewer/gui/canvases/canvas_window_ui.py:175` |
| `_on_save_canvas` | `sxm_viewer/gui/canvases/canvas_window.py:1506` | 2 | `sxm_viewer/gui/canvases/canvas_window_ui.py:173` |
| `_on_load_canvas` | `sxm_viewer/gui/canvases/canvas_window.py:1509` | 2 | `sxm_viewer/gui/canvases/canvas_window_ui.py:174` |
| `_on_sb_release` | `sxm_viewer/gui/canvases/detail_preview_canvas.py:3610` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:3394` |
| `_reset_colors` | `sxm_viewer/gui/canvases/molecular_overlay.py:703` | 4 | `sxm_viewer/gui/canvases/molecular_overlay.py:563` |
| `available_molecule_render_styles` | `sxm_viewer/gui/canvases/molecular_overlay.py:106` | 3 | `sxm_viewer/gui/canvases/molecular_overlay.py:498` |
| `is_reversed` | `sxm_viewer/gui/colormap_manager.py:49` | 2 | `sxm_viewer/gui/colormap_gallery.py:435` |
| `_on_profile_mode_toggled` | `sxm_viewer/gui/controllers/image_compare.py:1511` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1254` |
| `_on_canvas_press` | `sxm_viewer/gui/controllers/image_compare.py:1535` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1303` |
| `_on_canvas_motion` | `sxm_viewer/gui/controllers/image_compare.py:1538` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1304` |
| `_on_canvas_release` | `sxm_viewer/gui/controllers/image_compare.py:1541` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1305` |
| `_on_canvas_leave` | `sxm_viewer/gui/controllers/image_compare.py:1544` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1306` |
| `_export_comparison` | `sxm_viewer/gui/controllers/image_compare.py:1553` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1257` |
| `_reset_transform` | `sxm_viewer/gui/controllers/image_compare.py:1635` | 4 | `sxm_viewer/gui/controllers/image_compare.py:1226` |
| `_swap_slots` | `sxm_viewer/gui/controllers/image_compare.py:1640` | 3 | `sxm_viewer/gui/controllers/image_compare.py:1229` |
| `_icd_on_canvas_leave` | `sxm_viewer/gui/controllers/image_compare.py:1003` | 4 | `sxm_viewer/gui/controllers/image_compare.py:1545` |
| `focus_history_entry_with_shift` | `sxm_viewer/gui/controllers/quick_crop.py:407` | 4 | `sxm_viewer/gui/controllers/quick_crop.py:519` |
| `clear_history` | `sxm_viewer/gui/controllers/quick_crop.py:447` | 3 | `sxm_viewer/gui/main_window.py:1510` |
| `_on_error` | `sxm_viewer/gui/controllers/report.py:296` | 4 | `sxm_viewer/gui/controllers/report.py:264` |
| `_workspace_tick_format_x` | `sxm_viewer/gui/dialogs/image_adjust.py:440` | 2 | `sxm_viewer/gui/dialogs/image_adjust.py:195` |
| `_workspace_tick_format_y` | `sxm_viewer/gui/dialogs/image_adjust.py:443` | 2 | `sxm_viewer/gui/dialogs/image_adjust.py:196` |
| `_on_thread_finished` | `sxm_viewer/gui/dialogs/matrix_fit.py:566` | 2 | `sxm_viewer/gui/dialogs/matrix_fit.py:534` |
| `_on_taper_changed` | `sxm_viewer/gui/dialogs/periodic_noise.py:357` | 3 | `sxm_viewer/gui/dialogs/periodic_noise.py:146` |
| `_on_profile_list_context_menu` | `sxm_viewer/gui/dialogs/profile_dialog.py:783` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:449` |
| `_add_overlay_from_active` | `sxm_viewer/gui/dialogs/profile_dialog.py:2788` | 3 | `sxm_viewer/gui/dialogs/profile_dialog.py:464` |
| `_on_channel_combo_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4485` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:3937` |
| `_on_slice_slider_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4560` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:3942` |
| `_on_clear_background` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7699` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7353` |
| `_on_relative_toggled` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7918` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7198` |
| `_on_item_check_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7923` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7037` |
| `_on_point_label_release` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:8772` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7073` |
| `_on_mouse_move` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:9724` | 2 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7070` |
| `_on_minima_release` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:9767` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7069` |
| `_on_offset_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:9845` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7144` |
| `_fit_all` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10093` | 2 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7349` |
| `_on_fit_progress` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10131` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10109` |
| `_on_options_toggled` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10201` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7402` |
| `_on_clear_activity_log` | `sxm_viewer/gui/main_window.py:2536` | 4 | `sxm_viewer/gui/main_window.py:1011` |
| `_show_spectro_marker_legend` | `sxm_viewer/gui/main_window.py:2737` | 2 | `sxm_viewer/gui/main_window_toolbar.py:314` |
| `_on_hide_shortcuts_panel` | `sxm_viewer/gui/main_window.py:3170` | 2 | `sxm_viewer/gui/main_window_layout.py:586` |
| `_on_shortcuts_never_show_clicked` | `sxm_viewer/gui/main_window.py:3173` | 2 | `sxm_viewer/gui/main_window_layout.py:580` |
| `_on_show_shortcuts_requested` | `sxm_viewer/gui/main_window.py:3176` | 2 | `sxm_viewer/gui/main_window_toolbar.py:328` |
| `_on_autosave_timer` | `sxm_viewer/gui/main_window.py:3580` | 2 | `sxm_viewer/gui/main_window.py:1858` |
| `on_discard_recovery_snapshot` | `sxm_viewer/gui/main_window.py:3628` | 3 | `sxm_viewer/gui/main_window_layout.py:473` |
| `_on_toggle_layout_mode` | `sxm_viewer/gui/main_window.py:4714` | 3 | `sxm_viewer/gui/main_window_toolbar.py:333` |
| `open_folder_by_path` | `sxm_viewer/gui/main_window.py:4788` | 4 | `sxm_viewer/gui/main_window.py:1869` |
| `_on_frame_zoom_changed` | `sxm_viewer/gui/main_window.py:6764` | 4 | `sxm_viewer/gui/main_window.py:1036` |
| `on_load_session` | `sxm_viewer/gui/main_window.py:7184` | 3 | `sxm_viewer/gui/main_window_toolbar.py:80` |
| `on_open_collection` | `sxm_viewer/gui/main_window.py:7188` | 4 | `sxm_viewer/gui/main_window.py:933` |
| `on_choose_current_collection` | `sxm_viewer/gui/main_window.py:7215` | 3 | `sxm_viewer/gui/main_window.py:931` |
| `on_add_current_preview_to_collection` | `sxm_viewer/gui/main_window.py:7228` | 3 | `sxm_viewer/gui/main_window_toolbar.py:114` |
| `on_add_active_popup_to_collection` | `sxm_viewer/gui/main_window.py:7232` | 3 | `sxm_viewer/gui/main_window_toolbar.py:115` |

