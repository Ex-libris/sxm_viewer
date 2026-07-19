# Shim census - SXMGridViewer

A shim is a method whose whole body forwards elsewhere. Shims let logic leave the class without breaking callers, but they keep the method count up and keep the class as the discovery surface. Retiring a group means pointing its callers at the target module directly and deleting the shims.

`sxm_viewer/gui/main_window.py`

- total methods: **516**
- pure shims: **141** (27%)
- real logic: **375** (9193 lines)

## Shims by forwarding target

| Target | Shims | Methods |
|---|---|---|
| `viewer_thumb_ui` | 19 | _thumb_dimensions, _resize_thumbnail_scale, clear_thumbs, populate_thumbnails_for_channel, on_thumb_sort_changed +14 |
| `self` | 14 | _apply_dark_mode, _on_mode_button_clicked, on_open_spectro_browser, _on_hide_shortcuts_panel, _on_shortcuts_never_show_clicked +9 |
| `self.filter_controller` | 13 | _filter_action_label, _clone_filter_source_views, _filter_pipeline_label_from_steps, _thumbnail_filter_steps, _thumbnail_filter_label +8 |
| `main_window_spectro` | 11 | _update_spectro_stats_label, _header_extent, _display_extent, _spectros_near_thumb_pos, _open_single_spectro_popup +6 |
| `viewer_measurement` | 10 | _on_start_profile, _on_start_angle, _disable_profile_mode, _disable_angle_mode, _on_exit_profile_mode +5 |
| `spectro_loading` | 9 | ensure_spectros_loaded, _schedule_pending_spectro_load, _run_pending_spectro_load, _run_pending_spectro_load_async, _reload_spectros +4 |
| `viewer_export` | 8 | _collect_channel_exports, on_export_pngs, on_export_xyz_files, on_export_selected_same_view, on_export_stp_files +3 |
| `spectro_overrides` | 7 | _current_spectro_assignment_target_image_key, _spectro_override_signature, _resolve_spectro_override_targets, _refresh_spectro_assignment_overrides, _apply_spectro_assignment_override +2 |
| `virtual_copies` | 7 | _virtual_copy_source_anchor, _create_virtual_copy_from_popup_view, _create_virtual_copy_from_drag_payload, _create_virtual_channel_copies, _create_virtual_view_copy +2 |
| `self.recent_files_controller` | 5 | _refresh_recent_dirs_menu, _record_recent_dir, _refresh_recent_session_dirs_menu, _normalize_recent_session_history, _record_recent_session |
| `viewer_thumbnails` | 5 | _thumbnail_filter_signature, _downsample_for_thumbnail, _get_thumbnail_array, _thumbnail_data_key, _invalidate_thumbnail_cache |
| `viewer_loader` | 5 | _parse_header_datetime, _scan_spectros, hydrate_spectro_entry, hydrate_spectro_entries, refresh_spectro_manifest |
| `spectro_controller` | 5 | _choose_image_for_spec, _extent_center, _spec_within_extent, _spec_frame_offset_info, _match_spec_to_image_by_hint |
| `spec_mapping` | 5 | _map_spec_to_pixels, _apply_thumb_crop_to_coords, _map_spec_by_grid, _fallback_spec_coords, _map_spec_by_spec_extent |
| `viewer_preview` | 4 | _build_metadata_html, _build_single_channel_view, _on_preview_value, on_preview_cmap_changed |
| `main_window_layout` | 3 | _create_lower_controls, _apply_lower_control_theme, _create_shortcuts_panel |
| `self.session_controller` | 3 | on_save_session_as, on_save_session, on_load_session |
| `drift_animation` | 3 | _on_drift_correct, _on_create_animation, _show_alignment_preview |
| `main_window_toolbar` | 2 | _create_toolbar, _update_toolbar_actions |
| `spectro_overlays` | 1 | _show_spectro_marker_legend |
| `spectro_details` | 1 | _spectroscopy_metadata_lines |
| `self.collection_controller` | 1 | on_collection_help |

## Retirement candidates (32)

> Shims with **no** external `viewer.X` / `self.X` access and **no** string reference. Still verify each against call sites inside the owning file, then delete. Run the smoke test after - static analysis has missed a live call site here before.

| Shim | Forwards to |
|---|---|
| `_apply_dark_mode` | `self` |
| `_apply_spectro_scan_results` | `spectro_loading` |
| `_apply_thumb_crop_to_coords` | `spec_mapping` |
| `_clear_filter_for_paths` | `self.filter_controller` |
| `_create_lower_controls` | `main_window_layout` |
| `_create_shortcuts_panel` | `main_window_layout` |
| `_create_toolbar` | `main_window_toolbar` |
| `_create_virtual_channel_copies` | `virtual_copies` |
| `_create_virtual_copy_from_drag_payload` | `virtual_copies` |
| `_create_virtual_copy_from_popup_view` | `virtual_copies` |
| `_create_virtual_crop_view` | `virtual_copies` |
| `_current_spectro_focus_spec` | `spectro_overrides` |
| `_filter_badge_text` | `self.filter_controller` |
| `_map_spec_by_grid` | `spec_mapping` |
| `_map_spec_by_spec_extent` | `spec_mapping` |
| `_on_autosave_timer` | `self` |
| `_on_create_animation` | `drift_animation` |
| `_on_drift_correct` | `drift_animation` |
| `_on_spectro_manifest_save_finished` | `spectro_loading` |
| `_open_custom_filter_dialog` | `self.filter_controller` |
| `_refresh_spectro_assignment_overrides` | `spectro_overrides` |
| `_resolve_spectro_override_targets` | `spectro_overrides` |
| `_run_pending_spectro_load` | `spectro_loading` |
| `_show_alignment_preview` | `drift_animation` |
| `_spectro_override_signature` | `spectro_overrides` |
| `_spectroscopy_metadata_lines` | `spectro_details` |
| `_virtual_copy_source_anchor` | `virtual_copies` |
| `on_dark_mode_toggled` | `self` |
| `on_export_stp_files` | `viewer_export` |
| `on_save_session_as` | `self.session_controller` |
| `refresh_spectro_manifest` | `viewer_loader` |
| `toggle_star_for_paths` | `viewer_thumb_ui` |

### Reached ONLY by string/getattr - never delete blindly

| Shim | Referenced as a string in |
|---|---|
| `_clear_spectro_thumb_multi_selection` | thumbnail_ui.py |
| `_on_clear_profile_measurement` | measurement.py |
| `_on_exit_profile_mode` | measurement.py |
| `_on_preview_value` | preview.py |
| `_on_show_profile_window` | measurement.py |
| `_on_start_angle` | measurement.py |
| `_on_start_profile` | measurement.py |
| `_record_recent_session` | session.py |
| `_resize_thumbnail_scale` | thumbnail_ui.py |
| `_spectros_near_thumb_pos` | overlays.py |
| `on_export_selected_same_view` | export.py |
| `on_preview_cmap_changed` | preview.py |
| `on_thumb_cmap_changed` | thumbnail_ui.py |
| `on_thumb_filter_changed` | thumbnail_ui.py |
| `on_thumb_sort_changed` | thumbnail_ui.py |

## Largest remaining real-logic methods

| Lines | Line | Method |
|---|---|---|
| 1143 | 361 | `__init__` |
| 228 | 8931 | `_on_thumb_context_menu` |
| 129 | 10034 | `_apply_canvas_display_options` |
| 120 | 2861 | `_restore_closed_window_payload` |
| 119 | 2401 | `_apply_canvas_style_snapshot` |
| 114 | 7670 | `copy_selected_as_svg` |
| 112 | 3319 | `eventFilter` |
| 107 | 1676 | `_apply_ui_theme` |
| 107 | 8070 | `_on_let_the_robot_clicked` |
| 105 | 4850 | `set_plot_typography` |
| 98 | 4718 | `clear_loaded_images` |
| 98 | 7342 | `_send_thumbnail_targets_to_powerpoint` |
| 91 | 5026 | `_decorate_thumbnail_pixmap` |
| 77 | 5298 | `_sync_view_cmaps_from_canvas` |
| 76 | 7824 | `_show_toast` |
| 75 | 1918 | `_apply_preview_workspace_theme` |
| 75 | 7441 | `on_adjust_image` |
| 73 | 8781 | `_handle_spec_hover` |
| 72 | 8708 | `_handle_spec_marker_click` |
| 68 | 3533 | `_rebalance_main_splitter` |
| 68 | 3846 | `_create_session_activity_strip` |
| 67 | 3915 | `_set_session_activity` |
| 67 | 5443 | `_refresh_thumbnail_pixmaps_for_paths` |
| 65 | 4491 | `_populate_browse_molecules_menu` |
| 64 | 2249 | `_open_matrix_explorer_for_file` |
| 60 | 2340 | `_capture_canvas_style_snapshot` |
| 60 | 4077 | `_rebuild_popup_menu` |
| 59 | 4220 | `_apply_layout_mode` |
| 58 | 3246 | `_handle_local_file_mime_drop` |
| 57 | 5880 | `on_relative_axes_toggled` |
| 57 | 9160 | `_on_spectro_thumb_context_menu` |
| 54 | 5541 | `_set_thumbnail_entry_cmap` |
| 53 | 6712 | `_refresh_collection_toolbar_menu_labels` |
| 52 | 8863 | `_highlight_spectrum_entry` |
| 51 | 9556 | `_detach_preview` |
| 50 | 4138 | `_refresh_popup_ui` |
| 50 | 7619 | `render_and_save_file_using_config` |
| 49 | 6986 | `on_add_view` |
| 48 | 6663 | `_refresh_collection_ui` |
| 47 | 2579 | `_on_preview_crop` |

