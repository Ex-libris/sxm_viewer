# Shim census - SXMGridViewer

A shim is a method whose whole body forwards elsewhere. Shims let logic leave the class without breaking callers, but they keep the method count up and keep the class as the discovery surface. Retiring a group means pointing its callers at the target module directly and deleting the shims.

`sxm_viewer/gui/main_window.py`

- total methods: **536**
- pure shims: **153** (29%)
- real logic: **383** (10094 lines)

## Shims by forwarding target

| Target | Shims | Methods |
|---|---|---|
| `self.filter_controller` | 29 | _apply_filters_to_array, _filter_action_label, _clone_filter_source_views, _normalize_preview_filter_steps, _filter_pipeline_label_from_steps +24 |
| `viewer_thumb_ui` | 19 | _thumb_dimensions, _resize_thumbnail_scale, clear_thumbs, populate_thumbnails_for_channel, on_thumb_sort_changed +14 |
| `self` | 16 | _apply_dark_mode, _on_mode_button_clicked, on_open_spectro_browser, _on_hide_shortcuts_panel, _on_shortcuts_never_show_clicked +11 |
| `main_window_spectro` | 11 | _update_spectro_stats_label, _header_extent, _display_extent, _spectros_near_thumb_pos, _open_single_spectro_popup +6 |
| `viewer_measurement` | 10 | _on_start_profile, _on_start_angle, _disable_profile_mode, _disable_angle_mode, _on_exit_profile_mode +5 |
| `self.recent_files_controller` | 9 | _refresh_recent_dirs_menu, _record_recent_dir, _clear_recent_dirs, _refresh_recent_session_dirs_menu, _normalize_recent_session_history +4 |
| `spectro_loading` | 9 | ensure_spectros_loaded, _schedule_pending_spectro_load, _run_pending_spectro_load, _run_pending_spectro_load_async, _reload_spectros +4 |
| `viewer_export` | 8 | _collect_channel_exports, on_export_pngs, on_export_xyz_files, on_export_selected_same_view, on_export_stp_files +3 |
| `spectro_overrides` | 7 | _current_spectro_assignment_target_image_key, _spectro_override_signature, _resolve_spectro_override_targets, _refresh_spectro_assignment_overrides, _apply_spectro_assignment_override +2 |
| `viewer_thumbnails` | 5 | _thumbnail_filter_signature, _downsample_for_thumbnail, _get_thumbnail_array, _thumbnail_data_key, _invalidate_thumbnail_cache |
| `viewer_loader` | 5 | _parse_header_datetime, _scan_spectros, hydrate_spectro_entry, hydrate_spectro_entries, refresh_spectro_manifest |
| `spectro_controller` | 5 | _choose_image_for_spec, _extent_center, _spec_within_extent, _spec_frame_offset_info, _match_spec_to_image_by_hint |
| `spec_mapping` | 5 | _map_spec_to_pixels, _apply_thumb_crop_to_coords, _map_spec_by_grid, _fallback_spec_coords, _map_spec_by_spec_extent |
| `viewer_preview` | 4 | _build_metadata_html, _build_single_channel_view, _on_preview_value, on_preview_cmap_changed |
| `main_window_layout` | 3 | _create_lower_controls, _apply_lower_control_theme, _create_shortcuts_panel |
| `self.session_controller` | 3 | on_save_session_as, on_save_session, on_load_session |
| `main_window_toolbar` | 2 | _create_toolbar, _update_toolbar_actions |
| `spectro_overlays` | 1 | _show_spectro_marker_legend |
| `spectro_details` | 1 | _spectroscopy_metadata_lines |
| `self.collection_controller` | 1 | on_collection_help |

## Largest remaining real-logic methods

| Lines | Line | Method |
|---|---|---|
| 1143 | 359 | `__init__` |
| 277 | 9484 | `_on_create_animation` |
| 271 | 9762 | `_show_alignment_preview` |
| 228 | 9003 | `_on_thumb_context_menu` |
| 135 | 9348 | `_on_drift_correct` |
| 129 | 10991 | `_apply_canvas_display_options` |
| 120 | 2859 | `_restore_closed_window_payload` |
| 119 | 2399 | `_apply_canvas_style_snapshot` |
| 116 | 10131 | `_create_virtual_view_copy` |
| 114 | 7742 | `copy_selected_as_svg` |
| 112 | 3317 | `eventFilter` |
| 107 | 1674 | `_apply_ui_theme` |
| 107 | 8142 | `_on_let_the_robot_clicked` |
| 105 | 4855 | `set_plot_typography` |
| 98 | 4723 | `clear_loaded_images` |
| 98 | 7414 | `_send_thumbnail_targets_to_powerpoint` |
| 91 | 5031 | `_decorate_thumbnail_pixmap` |
| 77 | 5303 | `_sync_view_cmaps_from_canvas` |
| 76 | 7896 | `_show_toast` |
| 75 | 1916 | `_apply_preview_workspace_theme` |
| 75 | 7513 | `on_adjust_image` |
| 73 | 8853 | `_handle_spec_hover` |
| 72 | 8780 | `_handle_spec_marker_click` |
| 68 | 3531 | `_rebalance_main_splitter` |
| 68 | 3844 | `_create_session_activity_strip` |
| 67 | 3913 | `_set_session_activity` |
| 67 | 5448 | `_refresh_thumbnail_pixmaps_for_paths` |
| 65 | 4496 | `_populate_browse_molecules_menu` |
| 64 | 2247 | `_open_matrix_explorer_for_file` |
| 62 | 10068 | `_create_virtual_channel_copies` |
| 60 | 2338 | `_capture_canvas_style_snapshot` |
| 60 | 4075 | `_rebuild_popup_menu` |
| 59 | 4218 | `_apply_layout_mode` |
| 58 | 3244 | `_handle_local_file_mime_drop` |
| 57 | 5885 | `on_relative_axes_toggled` |
| 57 | 9232 | `_on_spectro_thumb_context_menu` |
| 54 | 5546 | `_set_thumbnail_entry_cmap` |
| 53 | 6720 | `_refresh_collection_toolbar_menu_labels` |
| 52 | 8935 | `_highlight_spectrum_entry` |
| 51 | 10513 | `_detach_preview` |

