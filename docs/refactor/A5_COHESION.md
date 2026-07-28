# A5 - Attribute cohesion (hidden class boundaries)

Clusters a large class's attributes by which methods use them. Attributes that always travel together are one piece of state - a candidate extracted class. Isolation scores rank which groups can be pulled out with the least ripple. Read-only evidence; no refactoring is implied by inclusion here.

## SXMGridViewer

`sxm_viewer/gui/main_window.py`

- methods analysed: **537**
- distinct `self.` attributes: **814**
- cohesive groups (>= 3 attrs): **30**

### Candidate extractions, most isolated first

`Isolation` = share of methods touching this group that touch *nothing else*. High = safe to extract.

| Isolation | Attrs | Pure/Touching methods | Attributes |
|---|---|---|---|
| 75% | 3 | 3/4 | `_activity_log_pending`, `activity_log_box`, `_activity_log_flush_timer` |
| 50% | 4 | 1/2 | `_session_activity_strip`, `_session_activity_title`, `_session_activity_detail`, `_session_activity_progress` |
| 50% | 3 | 4/8 | `_suspend_window_history`, `_capture_window_state_payload`, `_push_closed_window_history` |
| 50% | 3 | 2/4 | `spec_folder_le`, `spec_folder_path`, `_set_spec_folder` |
| 40% | 3 | 2/5 | `MODE_BROWSE`, `MODE_MEASURE`, `MODE_SPECTRO` |
| 40% | 3 | 2/5 | `ui_theme`, `_sync_forced_cmap`, `_apply_ui_theme` |
| 33% | 4 | 1/3 | `_preview_request_timer`, `_pending_preview_request`, `_flush_preview_request`, `_preview_render_in_progress` |
| 33% | 4 | 1/3 | `on_recall_popouts`, `on_close_popouts`, `on_minimize_popouts`, `on_arrange_popouts` |
| 33% | 3 | 1/3 | `_thumbnail_render_state_pending_paths`, `_thumbnail_render_state_timer`, `_flush_thumbnail_render_state_refresh` |
| 25% | 3 | 1/4 | `auto_detect_tags`, `_workspace_loading`, `_auto_detect_tags_for_folder` |
| 25% | 3 | 1/4 | `open_spectro_browser`, `_set_spectro_browser_filters`, `_spec_matches_image_key` |
| 20% | 5 | 1/5 | `spectros_by_image`, `files_with_matrix`, `_spec_extent_cache`, `matrix_datasets`, `files_with_spectra` |
| 14% | 6 | 1/7 | `main_splitter`, `_update_preview_detach_button`, `_layout_sizes`, `_preview_panel`, `_preview_dialog`, `preview_detached` |
| 14% | 4 | 1/7 | `_get_adjust_spec`, `_adjustment_undo_stack`, `_refresh_adjusted_channel`, `_set_adjust_spec` |
| 12% | 8 | 1/8 | `_thumb_inflight`, `_thumb_loaded`, `_request_visible_thumbs`, `_thumb_generation`, `_thumb_crop_cache`, `_thumb_labels` +2 |
| 0% | 19 | 0/4 | `thumbnail_filters`, `_deferred_popup_entries`, `spectro_groups_by_image`, `current_spectro_thumb_files`, `_update_toolbar_actions`, `spectro_sites_by_image` +13 |
| 0% | 7 | 0/6 | `show_molecules`, `_canvas_display_syncing`, `show_acquisition_overlay`, `scale_bar_cb`, `show_molecule_gizmo`, `_last_canvas_display_options` +1 |
| 0% | 6 | 0/3 | `_pending_compact_histogram_clim`, `_compact_histogram_gesture_active`, `_pending_compact_histogram_final`, `_suppress_compact_histogram_refresh`, `_compact_histogram_apply_timer`, `_flush_compact_histogram_clim` |
| 0% | 5 | 0/4 | `_highlighted_spec`, `_highlight_pulse_strength`, `_highlight_timer`, `_highlight_phase`, `_on_highlight_tick` |
| 0% | 4 | 0/4 | `on_show_molecules_toggled`, `_on_recent_molecules_updated`, `on_load_molecule`, `_on_molecule_palette_changed` |
| 0% | 4 | 0/4 | `starred`, `on_adjust_image`, `report_controller`, `on_add_selected_thumbnails_to_collection` |
| 0% | 4 | 0/5 | `thumb_cmap_combo`, `frame_real_view`, `thumb_cmap`, `_thumbnail_cmap_override` |
| 0% | 4 | 0/2 | `_plot_font_underline`, `_plot_font_family`, `_plot_font_italic`, `_plot_font_bold` |
| 0% | 3 | 0/6 | `_filtered_channel_cache`, `_frame_real_pixmap_cache`, `_filtered_cache_lock` |
| 0% | 3 | 0/4 | `spectro_thumb_channel_by_path`, `spectro_miniature_default_channel`, `spectro_compare_controller` |

### Group detail

<details><summary>3 attributes, 75% isolated (3/4 methods)</summary>

- `self._activity_log_flush_timer` - 2 methods
- `self._activity_log_pending` - 4 methods
- `self.activity_log_box` - 2 methods

</details>

<details><summary>4 attributes, 50% isolated (1/2 methods)</summary>

- `self._session_activity_detail` - 2 methods
- `self._session_activity_progress` - 2 methods
- `self._session_activity_strip` - 2 methods
- `self._session_activity_title` - 2 methods

</details>

<details><summary>3 attributes, 50% isolated (4/8 methods)</summary>

- `self._capture_window_state_payload` - 5 methods
- `self._push_closed_window_history` - 5 methods
- `self._suspend_window_history` - 8 methods

</details>

<details><summary>3 attributes, 50% isolated (2/4 methods)</summary>

- `self._set_spec_folder` - 2 methods
- `self.spec_folder_le` - 3 methods
- `self.spec_folder_path` - 3 methods

</details>

<details><summary>3 attributes, 40% isolated (2/5 methods)</summary>

- `self.MODE_BROWSE` - 5 methods
- `self.MODE_MEASURE` - 4 methods
- `self.MODE_SPECTRO` - 4 methods

</details>

<details><summary>3 attributes, 40% isolated (2/5 methods)</summary>

- `self._apply_ui_theme` - 3 methods
- `self._sync_forced_cmap` - 3 methods
- `self.ui_theme` - 3 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self._flush_preview_request` - 3 methods
- `self._pending_preview_request` - 3 methods
- `self._preview_render_in_progress` - 2 methods
- `self._preview_request_timer` - 3 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self.on_arrange_popouts` - 2 methods
- `self.on_close_popouts` - 2 methods
- `self.on_minimize_popouts` - 2 methods
- `self.on_recall_popouts` - 3 methods

</details>

<details><summary>3 attributes, 33% isolated (1/3 methods)</summary>

- `self._flush_thumbnail_render_state_refresh` - 2 methods
- `self._thumbnail_render_state_pending_paths` - 3 methods
- `self._thumbnail_render_state_timer` - 2 methods

</details>

<details><summary>3 attributes, 25% isolated (1/4 methods)</summary>

- `self._auto_detect_tags_for_folder` - 3 methods
- `self._workspace_loading` - 3 methods
- `self.auto_detect_tags` - 4 methods

</details>

<details><summary>3 attributes, 25% isolated (1/4 methods)</summary>

- `self._set_spectro_browser_filters` - 2 methods
- `self._spec_matches_image_key` - 2 methods
- `self.open_spectro_browser` - 4 methods

</details>

<details><summary>5 attributes, 20% isolated (1/5 methods)</summary>

- `self._spec_extent_cache` - 3 methods
- `self.files_with_matrix` - 3 methods
- `self.files_with_spectra` - 3 methods
- `self.matrix_datasets` - 3 methods
- `self.spectros_by_image` - 5 methods

</details>

### Extraction manifests (most isolated groups)

`pure` methods touch only this group's attributes and can move wholesale. `mixed` methods touch the group *and* other state - they stay put and call into the extracted object.

<details><summary><b>3 attrs, 75% isolated</b> - `_activity_log_flush_timer`, `_activity_log_pending`, `activity_log_box`</summary>

**State to move (3):**
- `self._activity_log_flush_timer`
- `self._activity_log_pending`
- `self.activity_log_box`

**Pure methods (3) - move these:**
- `_append_activity_log()`
- `_flush_activity_log_pending()`
- `_on_clear_activity_log()`

**Mixed methods (1) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_active_preview_canvas`, `_active_preview_popup` ...

</details>

<details><summary><b>4 attrs, 50% isolated</b> - `_session_activity_detail`, `_session_activity_progress`, `_session_activity_strip`, `_session_activity_title`</summary>

**State to move (4):**
- `self._session_activity_detail`
- `self._session_activity_progress`
- `self._session_activity_strip`
- `self._session_activity_title`

**Pure methods (1) - move these:**
- `_create_session_activity_strip()`

**Mixed methods (1) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_active_preview_canvas`, `_active_preview_popup` ...

</details>

<details><summary><b>3 attrs, 50% isolated</b> - `_capture_window_state_payload`, `_push_closed_window_history`, `_suspend_window_history`</summary>

**State to move (3):**
- `self._capture_window_state_payload`
- `self._push_closed_window_history`
- `self._suspend_window_history`

**Pure methods (4) - move these:**
- `_remember_closed_canvas_window()`
- `_remember_closed_main_profile_dialog()`
- `_remember_closed_popup_profile_dialog()`
- `_remember_closed_spectro_dialog()`

**Mixed methods (4) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_active_preview_canvas`, `_active_preview_popup` ...
- `_close_workspace_windows()` - also touches `_iter_workspace_windows`
- `_push_closed_window_history()` - also touches `_closed_window_history`
- `_remember_closed_preview_popup()` - also touches `_window_history_views_dir`, `session_controller`

</details>

<details><summary><b>3 attrs, 50% isolated</b> - `_set_spec_folder`, `spec_folder_le`, `spec_folder_path`</summary>

**State to move (3):**
- `self._set_spec_folder`
- `self.spec_folder_le`
- `self.spec_folder_path`

**Pure methods (2) - move these:**
- `on_spec_folder_browse()`
- `on_spec_folder_entered()`

**Mixed methods (2) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_active_preview_canvas`, `_active_preview_popup` ...
- `_set_spec_folder()` - also touches `_reload_spectros`, `config`

</details>

<details><summary><b>3 attrs, 40% isolated</b> - `MODE_BROWSE`, `MODE_MEASURE`, `MODE_SPECTRO`</summary>

**State to move (3):**
- `self.MODE_BROWSE`
- `self.MODE_MEASURE`
- `self.MODE_SPECTRO`

**Pure methods (2) - move these:**
- `_mode_from_name()`
- `_mode_name()`

**Mixed methods (3) - keep, delegate:**
- `_apply_mode()` - also touches `_disable_profile_mode`, `_mode_name`, `_on_start_profile`, `current_mode`, `mode_stack`
- `_init_mode_shortcuts()` - also touches `_mode_shortcuts`, `_on_mode_shortcut`
- `eventFilter()` - also touches `_clear_spectro_thumb_multi_selection`, `_clear_thumb_multi_selection`, `_focus_widget_blocks_thumb_nav`, `_handle_thumbnail_drag_event`, `_handle_thumbnail_navigation` ...

</details>

<details><summary><b>3 attrs, 40% isolated</b> - `_apply_ui_theme`, `_sync_forced_cmap`, `ui_theme`</summary>

**State to move (3):**
- `self._apply_ui_theme`
- `self._sync_forced_cmap`
- `self.ui_theme`

**Pure methods (2) - move these:**
- `_apply_dark_mode()`
- `_sync_forced_cmap()`

**Mixed methods (3) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_active_preview_canvas`, `_active_preview_popup` ...
- `on_amber_full_imagery_toggled()` - also touches `_refresh_all_imagery`, `amber_full_imagery`, `config`
- `set_ui_theme()` - also touches `_refresh_all_imagery`, `_set_detail_dark_view_state`, `config`, `dark_mode`, `detail_dark_view` ...

</details>

<details><summary><b>4 attrs, 33% isolated</b> - `_flush_preview_request`, `_pending_preview_request`, `_preview_render_in_progress`, `_preview_request_timer`</summary>

**State to move (4):**
- `self._flush_preview_request`
- `self._pending_preview_request`
- `self._preview_render_in_progress`
- `self._preview_request_timer`

**Pure methods (1) - move these:**
- `request_show_file_channel()`

**Mixed methods (2) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_active_preview_canvas`, `_active_preview_popup` ...
- `_flush_preview_request()` - also touches `show_file_channel`

</details>

<details><summary><b>4 attrs, 33% isolated</b> - `on_arrange_popouts`, `on_close_popouts`, `on_minimize_popouts`, `on_recall_popouts`</summary>

**State to move (4):**
- `self.on_arrange_popouts`
- `self.on_close_popouts`
- `self.on_minimize_popouts`
- `self.on_recall_popouts`

**Pure methods (1) - move these:**
- `on_restore_popouts()`

**Mixed methods (2) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_active_preview_canvas`, `_active_preview_popup` ...
- `_rebuild_popup_menu()` - also touches `_active_popup_windows`, `_describe_deferred_popup_entry`, `_focus_popup_window`, `_popup_window_menu_label`, `fontMetrics` ...

</details>

### Spine attributes (touched by the most methods)

> These resist extraction and define what the class is *actually* about; everything else is a passenger.

| Attribute | Methods touching |
|---|---|
| `self.config` | 72 |
| `self.last_preview` | 44 |
| `self.channel_dropdown` | 38 |
| `self.show_file_channel` | 38 |
| `self.filter_controller` | 31 |
| `self.populate_thumbnails_for_channel` | 29 |
| `self._spectros_loaded` | 22 |
| `self.collection_controller` | 19 |
| `self.preview_canvas` | 18 |
| `self.headers` | 18 |
| `self._schedule_marker_refresh` | 16 |
| `self.ensure_spectros_loaded` | 14 |
| `self.session_controller` | 11 |
| `self._is_processed_key` | 11 |
| `self._processed_views` | 10 |

