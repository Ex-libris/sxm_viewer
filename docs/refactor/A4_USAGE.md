# A4 - Usage and reachability

Name-based call analysis. **Never called** is a review queue, not a delete list - Qt signal connections and `getattr` dispatch are invisible here (flagged where a matching string literal exists).

## True phantom calls (never defined, never assigned)

None found.

## Injected callbacks (45 names)

> Called as `self.x(...)` with no `def x`, but assigned somewhere as an attribute - a deliberate plug-in point. Not bugs, but an **implicit, untyped interface**: nothing declares them, so a typo or a rename fails silently at runtime. Worth documenting as a real extension surface.

| Callback | Call sites | First site |
|---|---|---|
| `_value_callback` | 4 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:11516` |
| `_label_scale_cb` | 4 | `sxm_viewer/gui/dialogs/profile_dialog.py:513` |
| `_profile_marker_callback` | 3 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:4764` |
| `_profile_highlight_cb` | 3 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:7007` |
| `_preview_callback` | 3 | `sxm_viewer/gui/dialogs/filters.py:364` |
| `_copy_feedback_handler` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:5072` |
| `_profile_state_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8162` |
| `_views_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:1211` |
| `_crop_callback` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:11633` |
| `_last_svg_molecule_dir_cb` | 2 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8458` |
| `apply_step` | 2 | `sxm_viewer/gui/dialogs/filters.py:351` |
| `_marker_update_cb` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:1835` |
| `_marker_key_cb` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:1042` |
| `_highlight_overlay_cb` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:2889` |
| `_usage_provider` | 1 | `sxm_viewer/cmap_sorting.py:301` |
| `angle_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:5053` |
| `_default_preset_getter` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:716` |
| `_default_preset_setter` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:731` |
| `_font_change_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:1644` |
| `profile_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8148` |
| `_svg_molecule_style_defaults_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8667` |
| `_molecule_palette_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9052` |
| `_ch_height_menu_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9983` |
| `_filter_menu_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9996` |
| `_fixed_crop_history_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:12452` |
| `_double_click_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8229` |
| `_spectra_click_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8277` |
| `_recent_svg_molecule_cb` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:8619` |
| `_display_relative_zero_menu_state_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:9965` |
| `_histogram_auto_callback` | 1 | `sxm_viewer/gui/canvases/detail_preview_canvas.py:5458` |

## Never referenced (102 definitions)

- 83 with no matching string literal (strongest delete candidates)
- 19 whose name appears as a string somewhere (likely dynamic use - verify before touching)

| Name | Location | Lines | Owner |
|---|---|---|---|
| `_fast_spec_load_data` | `sxm_viewer/providers/nanonis/adapter.py:339` | 29 | (module) |
| `_safe_resize_event` | `sxm_viewer/_shared.py:43` | 21 | (module) |
| `grouped_base_cmap_names` | `sxm_viewer/cmap_registry.py:396` | 18 | (module) |
| `grouped_cmap_names` | `sxm_viewer/cmap_registry.py:279` | 16 | (module) |
| `bond_pairs` | `sxm_viewer/gui/canvases/svg_molecule_overlay.py:1196` | 10 | SvgMoleculeOverlay |
| `_toggle_inset_label_field` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:504` | 10 | (module) |
| `_site_sort_tuple` | `sxm_viewer/gui/spectroscopy/controller.py:618` | 9 | (module) |
| `test_never_produces_the_1e_minus_15_symptom` | `sxm_viewer/tests/test_profile_units.py:62` | 9 | TestToSiBase |
| `_lj_value` | `sxm_viewer/gui/dialogs/potential_fits.py:56` | 8 | (module) |
| `test_value_and_tolerance_share_one_si_prefix` | `sxm_viewer/tests/test_profile_measure.py:20` | 8 | TestFmtTol |
| `test_basic_stats_over_a_linear_ramp` | `sxm_viewer/tests/test_profile_measure.py:68` | 8 | TestBandStatistics |
| `test_unit_is_case_insensitive_for_recognition_but_preserves_symbol` | `sxm_viewer/tests/test_profile_units.py:36` | 8 | TestToSiBase |
| `stage` | `sxm_viewer/utils/logging.py:17` | 8 | (module) |
| `_again_obsidian` | `sxm_viewer/gui/viewer/export_routes.py:70` | 7 | (module) |
| `test_never_claims_more_precision_than_a_bare_format` | `sxm_viewer/tests/test_profile_measure.py:29` | 7 | TestFmtTol |
| `test_ignores_peaks_below_fraction_threshold` | `sxm_viewer/tests/test_profile_measure.py:58` | 7 | TestFindPeaks |
| `test_handles_reversed_edges` | `sxm_viewer/tests/test_profile_measure.py:77` | 7 | TestBandStatistics |
| `test_single_point_band_omits_derived_stats` | `sxm_viewer/tests/test_profile_measure.py:85` | 7 | TestBandStatistics |
| `test_constant_signal_has_zero_slope_and_rms` | `sxm_viewer/tests/test_profile_measure.py:93` | 7 | TestBandStatistics |
| `_lj_force` | `sxm_viewer/gui/dialogs/potential_fits.py:66` | 6 | (module) |
| `_lj_force_model` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:6527` | 6 | (module) |
| `unregister_profile_canvas` | `sxm_viewer/gui/profile_links.py:35` | 6 | (module) |
| `test_finds_single_clean_peak` | `sxm_viewer/tests/test_profile_measure.py:43` | 6 | TestFindPeaks |
| `test_never_produces_a_bare_multiplier` | `sxm_viewer/tests/test_profile_units.py:155` | 6 | TestSiFormat |
| `hoverEnterEvent` | `sxm_viewer/gui/canvases/canvas_items.py:487` | 5 | CanvasImageItem |
| `_morse_force` | `sxm_viewer/gui/dialogs/potential_fits.py:87` | 5 | (module) |
| `_pick_inset_marker_outline_color` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:497` | 5 | (module) |
| `test_snaps_each_edge_independently_within_tolerance` | `sxm_viewer/tests/test_profile_measure.py:118` | 5 | TestSnapEdges |
| `test_edge_far_from_any_peak_is_unchanged` | `sxm_viewer/tests/test_profile_measure.py:124` | 5 | TestSnapEdges |
| `test_nm_converts_to_metres` | `sxm_viewer/tests/test_profile_units.py:23` | 5 | TestToSiBase |
| `test_already_base_units_pass_through_unscaled` | `sxm_viewer/tests/test_profile_units.py:30` | 5 | TestToSiBase |
| `test_unrecognized_unit_passes_through_without_auto_prefix` | `sxm_viewer/tests/test_profile_units.py:45` | 5 | TestToSiBase |
| `test_blank_unit` | `sxm_viewer/tests/test_profile_units.py:51` | 5 | TestToSiBase |
| `test_strips_bracketed_unit_already_in_name` | `sxm_viewer/tests/test_profile_units.py:88` | 5 | TestComposeAxisLabel |
| `test_never_produces_doubled_unit_rendering` | `sxm_viewer/tests/test_profile_units.py:99` | 5 | TestComposeAxisLabel |
| `test_negative_values_use_the_same_prefix_as_positive` | `sxm_viewer/tests/test_profile_units.py:142` | 5 | TestSiScale |
| `dump_metadata_json` | `sxm_viewer/cmap_sorting.py:242` | 4 | (module) |
| `_morse_force_model` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:6535` | 4 | (module) |
| `test_flat_signal_has_no_peaks_above_threshold` | `sxm_viewer/tests/test_profile_measure.py:50` | 4 | TestFindPeaks |
| `test_counts_only_peaks_inside_band` | `sxm_viewer/tests/test_profile_measure.py:103` | 4 | TestPeaksInRange |
| `test_result_is_always_ordered_low_high` | `sxm_viewer/tests/test_profile_measure.py:134` | 4 | TestSnapEdges |
| `test_none_values_returns_empty_array` | `sxm_viewer/tests/test_profile_units.py:57` | 4 | TestToSiBase |
| `test_inverts_nm_conversion` | `sxm_viewer/tests/test_profile_units.py:74` | 4 | TestFromSiBase |
| `test_strips_parenthesized_unit_already_in_name` | `sxm_viewer/tests/test_profile_units.py:94` | 4 | TestComposeAxisLabel |
| `test_leaves_unrelated_brackets_alone` | `sxm_viewer/tests/test_profile_units.py:105` | 4 | TestComposeAxisLabel |
| `test_blank_unit_leaves_name_untouched` | `sxm_viewer/tests/test_profile_units.py:110` | 4 | TestComposeAxisLabel |
| `test_order_one_value_gets_no_prefix` | `sxm_viewer/tests/test_profile_units.py:117` | 4 | TestSiScale |
| `test_pico_range_current` | `sxm_viewer/tests/test_profile_units.py:122` | 4 | TestSiScale |
| `test_nano_range_distance` | `sxm_viewer/tests/test_profile_units.py:127` | 4 | TestSiScale |
| `test_kilo_range` | `sxm_viewer/tests/test_profile_units.py:132` | 4 | TestSiScale |
| `test_zero_is_handled_without_raising` | `sxm_viewer/tests/test_profile_units.py:137` | 4 | TestSiScale |
| `test_formats_with_chosen_prefix_and_unit` | `sxm_viewer/tests/test_profile_units.py:150` | 4 | TestSiFormat |
| `test_blank_name` | `sxm_viewer/tests/test_profile_units.py:166` | 4 | TestSiFormat |
| `extra_cmap_names` | `sxm_viewer/cmap_registry.py:244` | 3 | (module) |
| `set_grid_size` | `sxm_viewer/gui/canvases/canvas_view.py:35` | 3 | CanvasGraphicsView |
| `set_reversed` | `sxm_viewer/gui/colormap_manager.py:68` | 3 | ColormapManager |
| `pending_cmap_name` | `sxm_viewer/gui/colormap_manager.py:112` | 3 | ColormapManager |
| `pending_colormap` | `sxm_viewer/gui/colormap_manager.py:119` | 3 | ColormapManager |
| `save_session_as` | `sxm_viewer/gui/controllers/session.py:202` | 3 | SessionController |
| `_morse_value` | `sxm_viewer/gui/dialogs/potential_fits.py:82` | 3 | (module) |

## Called exactly once (663 definitions)

> Inline candidates - strongest for short shims.

104 of these are <= 4 lines.

| Name | Defined | Lines | Only caller |
|---|---|---|---|
| `project_root` | `sxm_viewer/app_meta.py:27` | 2 | `sxm_viewer/app_meta.py:32` |
| `samples_dir` | `sxm_viewer/app_meta.py:31` | 2 | `sxm_viewer/app_meta.py:43` |
| `assets_dir` | `sxm_viewer/app_meta.py:35` | 3 | `sxm_viewer/app_meta.py:43` |
| `_split_tokens` | `sxm_viewer/data/spectroscopy.py:543` | 3 | `sxm_viewer/data/spectroscopy.py:95` |
| `_relative` | `sxm_viewer/figures/height_series.py:841` | 4 | `sxm_viewer/figures/height_series.py:1502` |
| `_placeholder` | `sxm_viewer/figures/height_series.py:1995` | 4 | `sxm_viewer/figures/height_series.py:1550` |
| `_draw_corner_text` | `sxm_viewer/figures/height_series.py:2028` | 4 | `sxm_viewer/figures/height_series.py:1633` |
| `data_array` | `sxm_viewer/gui/canvases/canvas_items.py:1411` | 2 | `sxm_viewer/gui/canvases/canvas_window.py:409` |
| `safe_default_filename` | `sxm_viewer/gui/canvases/canvas_rendering.py:74` | 4 | `sxm_viewer/gui/canvases/canvas_items.py:1388` |
| `canvas_items` | `sxm_viewer/gui/canvases/canvas_state.py:11` | 2 | `sxm_viewer/gui/canvases/canvas_state.py:20` |
| `selected_canvas_items` | `sxm_viewer/gui/canvases/canvas_state.py:15` | 2 | `sxm_viewer/gui/canvases/canvas_state.py:63` |
| `_on_export_image` | `sxm_viewer/gui/canvases/canvas_window.py:1452` | 2 | `sxm_viewer/gui/canvases/canvas_window_ui.py:127` |
| `_on_save_canvas` | `sxm_viewer/gui/canvases/canvas_window.py:1455` | 2 | `sxm_viewer/gui/canvases/canvas_window_ui.py:125` |
| `_on_load_canvas` | `sxm_viewer/gui/canvases/canvas_window.py:1458` | 2 | `sxm_viewer/gui/canvases/canvas_window_ui.py:126` |
| `_relative_luminance` | `sxm_viewer/gui/canvases/scale_bar_contrast.py:25` | 4 | `sxm_viewer/gui/canvases/scale_bar_contrast.py:136` |
| `is_reversed` | `sxm_viewer/gui/colormap_manager.py:49` | 2 | `sxm_viewer/gui/colormap_gallery.py:435` |
| `_on_profile_mode_toggled` | `sxm_viewer/gui/controllers/image_compare.py:1497` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1240` |
| `_on_canvas_press` | `sxm_viewer/gui/controllers/image_compare.py:1521` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1289` |
| `_on_canvas_motion` | `sxm_viewer/gui/controllers/image_compare.py:1524` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1290` |
| `_on_canvas_release` | `sxm_viewer/gui/controllers/image_compare.py:1527` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1291` |
| `_on_canvas_leave` | `sxm_viewer/gui/controllers/image_compare.py:1530` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1292` |
| `_export_comparison` | `sxm_viewer/gui/controllers/image_compare.py:1539` | 2 | `sxm_viewer/gui/controllers/image_compare.py:1243` |
| `_reset_transform` | `sxm_viewer/gui/controllers/image_compare.py:1621` | 4 | `sxm_viewer/gui/controllers/image_compare.py:1212` |
| `_swap_slots` | `sxm_viewer/gui/controllers/image_compare.py:1626` | 3 | `sxm_viewer/gui/controllers/image_compare.py:1215` |
| `_icd_on_canvas_leave` | `sxm_viewer/gui/controllers/image_compare.py:989` | 4 | `sxm_viewer/gui/controllers/image_compare.py:1531` |
| `focus_history_entry_with_shift` | `sxm_viewer/gui/controllers/quick_crop.py:407` | 4 | `sxm_viewer/gui/controllers/quick_crop.py:519` |
| `clear_history` | `sxm_viewer/gui/controllers/quick_crop.py:447` | 3 | `sxm_viewer/gui/main_window.py:1126` |
| `_on_error` | `sxm_viewer/gui/controllers/report.py:306` | 4 | `sxm_viewer/gui/controllers/report.py:274` |
| `_on_format_changed` | `sxm_viewer/gui/dialogs/figure_window.py:1492` | 4 | `sxm_viewer/gui/dialogs/figure_window.py:1439` |
| `_figure_classes` | `sxm_viewer/gui/dialogs/figure_window.py:134` | 4 | `sxm_viewer/gui/dialogs/figure_window.py:1645` |
| `_workspace_tick_format_x` | `sxm_viewer/gui/dialogs/image_adjust.py:347` | 2 | `sxm_viewer/gui/dialogs/image_adjust.py:102` |
| `_workspace_tick_format_y` | `sxm_viewer/gui/dialogs/image_adjust.py:350` | 2 | `sxm_viewer/gui/dialogs/image_adjust.py:103` |
| `_on_thread_finished` | `sxm_viewer/gui/dialogs/matrix_fit.py:519` | 2 | `sxm_viewer/gui/dialogs/matrix_fit.py:487` |
| `_on_taper_changed` | `sxm_viewer/gui/dialogs/periodic_noise.py:357` | 3 | `sxm_viewer/gui/dialogs/periodic_noise.py:146` |
| `_noise_sigma` | `sxm_viewer/gui/dialogs/potential_fits.py:139` | 4 | `sxm_viewer/gui/dialogs/potential_fits.py:158` |
| `_on_span_move` | `sxm_viewer/gui/dialogs/profile_dialog.py:1671` | 4 | `sxm_viewer/gui/dialogs/profile_dialog.py:668` |
| `_on_grid_toggled` | `sxm_viewer/gui/dialogs/profile_dialog.py:1881` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:604` |
| `_on_profile_list_context_menu` | `sxm_viewer/gui/dialogs/profile_dialog.py:2394` | 2 | `sxm_viewer/gui/dialogs/profile_dialog.py:792` |
| `_set_include_band_in_export` | `sxm_viewer/gui/dialogs/profile_dialog.py:2522` | 3 | `sxm_viewer/gui/dialogs/profile_dialog.py:2466` |
| `_add_overlay_from_active` | `sxm_viewer/gui/dialogs/profile_dialog.py:2966` | 3 | `sxm_viewer/gui/dialogs/profile_dialog.py:757` |
| `_set_inset_marker_filled` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1941` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:711` |
| `_set_inset_show_other_points` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1946` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:680` |
| `_set_inset_marker_size` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:1958` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:708` |
| `_on_channel_combo_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:5003` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4455` |
| `_on_slice_slider_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:5078` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:4460` |
| `_on_force_source_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:8807` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:8149` |
| `_on_relative_toggled` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:9055` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7875` |
| `_on_item_check_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:9060` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7722` |
| `_on_potential_model_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:9714` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:8033` |
| `_on_point_label_release` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:10249` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7759` |
| `_on_mouse_move` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:11190` | 2 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7756` |
| `_on_minima_release` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:11233` | 4 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7755` |
| `_on_offset_changed` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:11311` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:7821` |
| `_fit_all` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:11561` | 2 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:8217` |
| `_on_fit_progress` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:11599` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:11577` |
| `_on_options_toggled` | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:11669` | 3 | `sxm_viewer/gui/dialogs/spectroscopy_dialogs.py:8274` |
| `_append_activity_log` | `sxm_viewer/gui/main_window.py:2128` | 4 | `sxm_viewer/gui/main_window.py:1584` |
| `_show_spectro_marker_legend` | `sxm_viewer/gui/main_window.py:2309` | 2 | `sxm_viewer/gui/main_window_toolbar.py:345` |
| `_on_hide_shortcuts_panel` | `sxm_viewer/gui/main_window.py:2744` | 2 | `sxm_viewer/gui/main_window_layout.py:591` |
| `_on_shortcuts_never_show_clicked` | `sxm_viewer/gui/main_window.py:2747` | 2 | `sxm_viewer/gui/main_window_layout.py:585` |

