# Shim census - SXMGridViewer

A shim is a method whose whole body forwards elsewhere. Shims let logic leave the class without breaking callers, but they keep the method count up and keep the class as the discovery surface. Retiring a group means pointing its callers at the target module directly and deleting the shims.

`sxm_viewer/gui/main_window.py`

- total methods: **497**
- pure shims: **122** (25%)
- real logic: **375** (9193 lines)

## Shims by forwarding target

| Target | Shims | Methods |
|---|---|---|
| `viewer_thumb_ui` | 19 | _thumb_dimensions, _resize_thumbnail_scale, clear_thumbs, populate_thumbnails_for_channel, on_thumb_sort_changed +14 |
| `self.filter_controller` | 13 | _filter_action_label, _clone_filter_source_views, _filter_pipeline_label_from_steps, _thumbnail_filter_steps, _thumbnail_filter_label +8 |
| `self` | 12 | _on_mode_button_clicked, on_open_spectro_browser, _on_hide_shortcuts_panel, _on_shortcuts_never_show_clicked, _on_show_shortcuts_requested +7 |
| `main_window_spectro` | 11 | _update_spectro_stats_label, _header_extent, _display_extent, _spectros_near_thumb_pos, _open_single_spectro_popup +6 |
| `viewer_measurement` | 10 | _on_start_profile, _on_start_angle, _disable_profile_mode, _disable_angle_mode, _on_exit_profile_mode +5 |
| `viewer_export` | 8 | _collect_channel_exports, on_export_pngs, on_export_xyz_files, on_export_selected_same_view, on_export_stp_files +3 |
| `spectro_loading` | 6 | ensure_spectros_loaded, _schedule_pending_spectro_load, _run_pending_spectro_load_async, _reload_spectros, _schedule_spectro_manifest_save +1 |
| `self.recent_files_controller` | 5 | _refresh_recent_dirs_menu, _record_recent_dir, _refresh_recent_session_dirs_menu, _normalize_recent_session_history, _record_recent_session |
| `viewer_thumbnails` | 5 | _thumbnail_filter_signature, _downsample_for_thumbnail, _get_thumbnail_array, _thumbnail_data_key, _invalidate_thumbnail_cache |
| `viewer_loader` | 5 | _parse_header_datetime, _scan_spectros, hydrate_spectro_entry, hydrate_spectro_entries, refresh_spectro_manifest |
| `spectro_controller` | 5 | _choose_image_for_spec, _extent_center, _spec_within_extent, _spec_frame_offset_info, _match_spec_to_image_by_hint |
| `viewer_preview` | 4 | _build_metadata_html, _build_single_channel_view, _on_preview_value, on_preview_cmap_changed |
| `virtual_copies` | 4 | _create_virtual_copy_from_popup_view, _create_virtual_channel_copies, _create_virtual_view_copy, _create_virtual_copy_from_history |
| `main_window_layout` | 3 | _create_lower_controls, _apply_lower_control_theme, _create_shortcuts_panel |
| `self.session_controller` | 3 | on_save_session_as, on_save_session, on_load_session |
| `spectro_overrides` | 3 | _current_spectro_assignment_target_image_key, _apply_spectro_assignment_override, _clear_spectro_assignment_override |
| `main_window_toolbar` | 2 | _create_toolbar, _update_toolbar_actions |
| `spec_mapping` | 2 | _map_spec_to_pixels, _fallback_spec_coords |
| `spectro_overlays` | 1 | _show_spectro_marker_legend |
| `self.collection_controller` | 1 | on_collection_help |

## Retirement candidates (1)

> No caller by **any** of the three routes: external `viewer.X`/`self.X`, `self.X` inside the owning class, or a string literal (getattr dispatch). Still run the smoke test after deleting - static analysis has already missed a live call site on this branch.

| Shim | Forwards to |
|---|---|
| `on_save_session_as` | `self.session_controller` |

## Internal-only shims (11)

> Called only from other methods of this class. Retiring these means rewriting those call sites to use the target object directly, then deleting the shim.

| Shim | Forwards to | Called by |
|---|---|---|
| `_clear_filter_for_paths` | `self.filter_controller` | _on_thumb_context_menu |
| `_create_lower_controls` | `main_window_layout` | __init__ |
| `_create_shortcuts_panel` | `main_window_layout` | __init__ |
| `_create_toolbar` | `main_window_toolbar` | __init__ |
| `_create_virtual_channel_copies` | `virtual_copies` | _on_thumb_context_menu |
| `_filter_badge_text` | `self.filter_controller` | _decorate_thumbnail_pixmap |
| `_on_autosave_timer` | `self` | __init__ |
| `_open_custom_filter_dialog` | `self.filter_controller` | _on_thumb_context_menu |
| `on_export_stp_files` | `viewer_export` | _on_thumb_context_menu |
| `refresh_spectro_manifest` | `viewer_loader` | _assign_spectros_to_images |
| `toggle_star_for_paths` | `viewer_thumb_ui` | _on_thumb_context_menu, keyPressEvent |

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
| `on_export_selected_same_view` | export.py |
| `on_preview_cmap_changed` | preview.py |
| `on_thumb_cmap_changed` | thumbnail_ui.py |
| `on_thumb_filter_changed` | thumbnail_ui.py |
| `on_thumb_sort_changed` | thumbnail_ui.py |

## Largest remaining real-logic methods

| Lines | Line | Method |
|---|---|---|
| 1143 | 361 | `__init__` |
| 228 | 8895 | `_on_thumb_context_menu` |
| 129 | 9977 | `_apply_canvas_display_options` |
| 120 | 2857 | `_restore_closed_window_payload` |
| 119 | 2397 | `_apply_canvas_style_snapshot` |
| 114 | 7667 | `copy_selected_as_svg` |
| 112 | 3320 | `eventFilter` |
| 107 | 1672 | `_apply_ui_theme` |
| 107 | 8067 | `_on_let_the_robot_clicked` |
| 105 | 4847 | `set_plot_typography` |
| 98 | 4719 | `clear_loaded_images` |
| 98 | 7339 | `_send_thumbnail_targets_to_powerpoint` |
| 91 | 5023 | `_decorate_thumbnail_pixmap` |
| 77 | 5295 | `_sync_view_cmaps_from_canvas` |
| 76 | 7821 | `_show_toast` |
| 75 | 1914 | `_apply_preview_workspace_theme` |
| 75 | 7438 | `on_adjust_image` |
| 73 | 8745 | `_handle_spec_hover` |
| 72 | 8672 | `_handle_spec_marker_click` |
| 68 | 3534 | `_rebalance_main_splitter` |
| 68 | 3847 | `_create_session_activity_strip` |
| 67 | 3916 | `_set_session_activity` |
| 67 | 5440 | `_refresh_thumbnail_pixmaps_for_paths` |
| 65 | 4492 | `_populate_browse_molecules_menu` |
| 64 | 2245 | `_open_matrix_explorer_for_file` |
| 60 | 2336 | `_capture_canvas_style_snapshot` |
| 60 | 4078 | `_rebuild_popup_menu` |
| 59 | 4221 | `_apply_layout_mode` |
| 58 | 3247 | `_handle_local_file_mime_drop` |
| 57 | 5877 | `on_relative_axes_toggled` |
| 57 | 9124 | `_on_spectro_thumb_context_menu` |
| 54 | 5538 | `_set_thumbnail_entry_cmap` |
| 53 | 6709 | `_refresh_collection_toolbar_menu_labels` |
| 52 | 8827 | `_highlight_spectrum_entry` |
| 51 | 9503 | `_detach_preview` |
| 50 | 4139 | `_refresh_popup_ui` |
| 50 | 7616 | `render_and_save_file_using_config` |
| 49 | 6983 | `on_add_view` |
| 48 | 6660 | `_refresh_collection_ui` |
| 47 | 2575 | `_on_preview_crop` |

