# A5 - Attribute cohesion (hidden class boundaries)

Clusters a large class's attributes by which methods use them. Attributes that always travel together are one piece of state - a candidate extracted class. Isolation scores rank which groups can be pulled out with the least ripple. Read-only evidence; no refactoring is implied by inclusion here.

## SXMGridViewer

`sxm_viewer/gui/main_window.py`

- methods analysed: **520**
- distinct `self.` attributes: **714**
- cohesive groups (>= 3 attrs): **20**

### Candidate extractions, most isolated first

`Isolation` = share of methods touching this group that touch *nothing else*. High = safe to extract.

| Isolation | Attrs | Pure/Touching methods | Attributes |
|---|---|---|---|
| 57% | 3 | 4/7 | `_suspend_window_history`, `_capture_window_state_payload`, `_push_closed_window_history` |
| 40% | 3 | 2/5 | `MODE_BROWSE`, `MODE_MEASURE`, `MODE_SPECTRO` |
| 33% | 4 | 1/3 | `on_recall_popouts`, `on_minimize_popouts`, `on_arrange_popouts`, `on_close_popouts` |
| 25% | 6 | 1/4 | `spectros_by_image`, `matrix_spectros`, `matrix_datasets`, `_spec_extent_cache`, `files_with_spectra`, `files_with_matrix` |
| 25% | 4 | 1/4 | `open_spectro_browser`, `_spec_matches_image_key`, `_set_spectro_browser_filters`, `_filter_spectro_browser` |
| 25% | 3 | 1/4 | `auto_detect_tags`, `_auto_detect_tags_for_folder`, `_workspace_loading` |
| 0% | 5 | 0/4 | `meta_box`, `_update_toolbar_actions`, `angle_value_label`, `preview_value_label`, `_refresh_deferred_popup_ui` |
| 0% | 4 | 0/7 | `main_splitter`, `_update_preview_detach_button`, `_layout_sizes`, `_preview_panel` |
| 0% | 4 | 0/4 | `thumb_cmap_combo`, `thumb_cmap`, `on_thumb_cmap_changed`, `_thumbnail_cmap_override` |
| 0% | 4 | 0/4 | `on_show_molecules_toggled`, `_on_recent_molecules_updated`, `on_load_molecule`, `_on_molecule_palette_changed` |
| 0% | 4 | 0/6 | `show_molecules`, `scale_bar_cb`, `show_acquisition_overlay`, `show_filter_summary` |
| 0% | 4 | 0/3 | `_highlighted_spec`, `_highlight_timer`, `_highlight_phase`, `_highlight_pulse_strength` |
| 0% | 4 | 0/2 | `_plot_font_family`, `_plot_font_underline`, `_plot_font_bold`, `_plot_font_italic` |
| 0% | 4 | 0/2 | `rotate_images`, `toggle_zero_images`, `_auto_contrast`, `set_level_filter` |
| 0% | 3 | 0/6 | `_thumb_inflight`, `_thumb_generation`, `_thumb_loaded` |
| 0% | 3 | 0/4 | `relative_axes`, `show_molecule_gizmo`, `canvas_display_options` |
| 0% | 3 | 0/2 | `_right_splitter`, `_right_container`, `_thumbs_panel` |
| 0% | 3 | 0/2 | `_handle_thumbnail_navigation`, `_select_all_thumbnails`, `_focus_widget_blocks_thumb_nav` |
| 0% | 3 | 0/2 | `_thumb_data_lock`, `_thumb_crop_cache`, `_decorate_thumbnail_pixmap` |
| 0% | 3 | 0/2 | `_pending_compact_histogram_clim`, `_pending_compact_histogram_final`, `_compact_histogram_gesture_active` |

### Group detail

<details><summary>3 attributes, 57% isolated (4/7 methods)</summary>

- `self._capture_window_state_payload` - 5 methods
- `self._push_closed_window_history` - 5 methods
- `self._suspend_window_history` - 7 methods

</details>

<details><summary>3 attributes, 40% isolated (2/5 methods)</summary>

- `self.MODE_BROWSE` - 5 methods
- `self.MODE_MEASURE` - 4 methods
- `self.MODE_SPECTRO` - 4 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self.on_arrange_popouts` - 2 methods
- `self.on_close_popouts` - 2 methods
- `self.on_minimize_popouts` - 2 methods
- `self.on_recall_popouts` - 3 methods

</details>

<details><summary>6 attributes, 25% isolated (1/4 methods)</summary>

- `self._spec_extent_cache` - 2 methods
- `self.files_with_matrix` - 2 methods
- `self.files_with_spectra` - 2 methods
- `self.matrix_datasets` - 2 methods
- `self.matrix_spectros` - 2 methods
- `self.spectros_by_image` - 4 methods

</details>

<details><summary>4 attributes, 25% isolated (1/4 methods)</summary>

- `self._filter_spectro_browser` - 2 methods
- `self._set_spectro_browser_filters` - 2 methods
- `self._spec_matches_image_key` - 2 methods
- `self.open_spectro_browser` - 4 methods

</details>

<details><summary>3 attributes, 25% isolated (1/4 methods)</summary>

- `self._auto_detect_tags_for_folder` - 3 methods
- `self._workspace_loading` - 2 methods
- `self.auto_detect_tags` - 4 methods

</details>

<details><summary>5 attributes, 0% isolated (0/4 methods)</summary>

- `self._refresh_deferred_popup_ui` - 2 methods
- `self._update_toolbar_actions` - 2 methods
- `self.angle_value_label` - 2 methods
- `self.meta_box` - 4 methods
- `self.preview_value_label` - 2 methods

</details>

<details><summary>4 attributes, 0% isolated (0/7 methods)</summary>

- `self._layout_sizes` - 4 methods
- `self._preview_panel` - 4 methods
- `self._update_preview_detach_button` - 4 methods
- `self.main_splitter` - 5 methods

</details>

<details><summary>4 attributes, 0% isolated (0/4 methods)</summary>

- `self._thumbnail_cmap_override` - 2 methods
- `self.on_thumb_cmap_changed` - 2 methods
- `self.thumb_cmap` - 3 methods
- `self.thumb_cmap_combo` - 4 methods

</details>

<details><summary>4 attributes, 0% isolated (0/4 methods)</summary>

- `self._on_molecule_palette_changed` - 2 methods
- `self._on_recent_molecules_updated` - 3 methods
- `self.on_load_molecule` - 2 methods
- `self.on_show_molecules_toggled` - 4 methods

</details>

<details><summary>4 attributes, 0% isolated (0/6 methods)</summary>

- `self.scale_bar_cb` - 3 methods
- `self.show_acquisition_overlay` - 3 methods
- `self.show_filter_summary` - 3 methods
- `self.show_molecules` - 3 methods

</details>

<details><summary>4 attributes, 0% isolated (0/3 methods)</summary>

- `self._highlight_phase` - 2 methods
- `self._highlight_pulse_strength` - 2 methods
- `self._highlight_timer` - 2 methods
- `self._highlighted_spec` - 3 methods

</details>

### Extraction manifests (most isolated groups)

`pure` methods touch only this group's attributes and can move wholesale. `mixed` methods touch the group *and* other state - they stay put and call into the extracted object.

<details><summary><b>3 attrs, 57% isolated</b> - `_capture_window_state_payload`, `_push_closed_window_history`, `_suspend_window_history`</summary>

**State to move (3):**
- `self._capture_window_state_payload`
- `self._push_closed_window_history`
- `self._suspend_window_history`

**Pure methods (4) - move these:**
- `_remember_closed_canvas_window()`
- `_remember_closed_main_profile_dialog()`
- `_remember_closed_popup_profile_dialog()`
- `_remember_closed_spectro_dialog()`

**Mixed methods (3) - keep, delegate:**
- `_close_workspace_windows()` - also touches `_iter_workspace_windows`
- `_push_closed_window_history()` - also touches `_closed_window_history`
- `_remember_closed_preview_popup()` - also touches `_window_history_views_dir`, `session_controller`

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

<details><summary><b>4 attrs, 33% isolated</b> - `on_arrange_popouts`, `on_close_popouts`, `on_minimize_popouts`, `on_recall_popouts`</summary>

**State to move (4):**
- `self.on_arrange_popouts`
- `self.on_close_popouts`
- `self.on_minimize_popouts`
- `self.on_recall_popouts`

**Pure methods (1) - move these:**
- `on_restore_popouts()`

**Mixed methods (2) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_app_start_ts`, `_append_activity_log` ...
- `_rebuild_popup_menu()` - also touches `_active_popup_windows`, `_describe_deferred_popup_entry`, `_focus_popup_window`, `_popup_window_menu_label`, `fontMetrics` ...

</details>

<details><summary><b>6 attrs, 25% isolated</b> - `_spec_extent_cache`, `files_with_matrix`, `files_with_spectra`, `matrix_datasets` ...</summary>

**State to move (6):**
- `self._spec_extent_cache`
- `self.files_with_matrix`
- `self.files_with_spectra`
- `self.matrix_datasets`
- `self.matrix_spectros`
- `self.spectros_by_image`

**Pure methods (1) - move these:**
- `_spec_cloud_bounds()`

**Mixed methods (3) - keep, delegate:**
- `_assign_spectros_to_images()` - also touches `_update_matrix_summary_banner`, `refresh_spectro_manifest`
- `_open_matrix_explorer_for_file()` - also touches `_popup_refs`, `_spectros_loaded`, `ensure_spectros_loaded`, `hydrate_spectro_entries`
- `clear_loaded_images()` - also touches `_collection_item_snapshots`, `_current_session_path`, `_deferred_popup_entries`, `_deferred_popup_serial`, `_frame_real_pixmap_cache` ...

</details>

<details><summary><b>4 attrs, 25% isolated</b> - `_filter_spectro_browser`, `_set_spectro_browser_filters`, `_spec_matches_image_key`, `open_spectro_browser`</summary>

**State to move (4):**
- `self._filter_spectro_browser`
- `self._set_spectro_browser_filters`
- `self._spec_matches_image_key`
- `self.open_spectro_browser`

**Pure methods (1) - move these:**
- `on_open_spectro_browser()`

**Mixed methods (3) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_app_start_ts`, `_append_activity_log` ...
- `on_review_low_conf_spectros()` - also touches `_current_spectro_assignment_target_image_key`, `_spectros_loaded`, `ensure_spectros_loaded`
- `on_review_off_frame_spectros()` - also touches `_current_spectro_assignment_target_image_key`, `_spectros_loaded`, `ensure_spectros_loaded`

</details>

<details><summary><b>3 attrs, 25% isolated</b> - `_auto_detect_tags_for_folder`, `_workspace_loading`, `auto_detect_tags`</summary>

**State to move (3):**
- `self._auto_detect_tags_for_folder`
- `self._workspace_loading`
- `self.auto_detect_tags`

**Pure methods (1) - move these:**
- `load_files()`

**Mixed methods (3) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_app_start_ts`, `_append_activity_log` ...
- `_on_toggle_auto_tags()` - also touches `_clear_auto_tags`, `channel_dropdown`, `config`, `last_preview`, `populate_thumbnails_for_channel` ...
- `load_folder()` - also touches `_notify_folder_collection_usage`, `_prepare_for_workspace_load`

</details>

<details><summary><b>5 attrs, 0% isolated</b> - `_refresh_deferred_popup_ui`, `_update_toolbar_actions`, `angle_value_label`, `meta_box` ...</summary>

**State to move (5):**
- `self._refresh_deferred_popup_ui`
- `self._update_toolbar_actions`
- `self.angle_value_label`
- `self.meta_box`
- `self.preview_value_label`

**Pure methods (0) - move these:**

**Mixed methods (4) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_app_start_ts`, `_append_activity_log` ...
- `clear_loaded_images()` - also touches `_collection_item_snapshots`, `_current_session_path`, `_deferred_popup_entries`, `_deferred_popup_serial`, `_frame_real_pixmap_cache` ...
- `on_meta_font_changed()` - also touches `_build_metadata_html`, `config`
- `show_spectroscopy_details()` - also touches `details_group`

</details>

<details><summary><b>4 attrs, 0% isolated</b> - `_layout_sizes`, `_preview_panel`, `_update_preview_detach_button`, `main_splitter`</summary>

**State to move (4):**
- `self._layout_sizes`
- `self._preview_panel`
- `self._update_preview_detach_button`
- `self.main_splitter`

**Pure methods (0) - move these:**

**Mixed methods (7) - keep, delegate:**
- `__init__()` - also touches `FRAME_ZOOM_SLIDER_DEFAULT`, `FRAME_ZOOM_SLIDER_MAX`, `FRAME_ZOOM_SLIDER_MIN`, `_app_start_ts`, `_append_activity_log` ...
- `_apply_layout_mode()` - also touches `_layout_mode`, `_left_sidebar_rebalance_timer`, `_right_container`, `_right_splitter`, `_thumbs_panel` ...
- `_attach_preview()` - also touches `_preview_dialog`, `preview_detached`
- `_detach_preview()` - also touches `_attach_preview`, `_preview_dialog`, `preview_detached`, `preview_locked`
- `_on_main_splitter_moved()` - also touches `_layout_mode`
- `_rebalance_main_splitter()` - also touches `_layout_mode`, `_scale_splitter_sizes`, `isMaximized`, `width`
- `on_preview_lock_toggled()` - also touches `_attach_preview`, `config`, `preview_locked`

</details>

### Spine attributes (touched by the most methods)

> These resist extraction and define what the class is *actually* about; everything else is a passenger.

| Attribute | Methods touching |
|---|---|
| `self.config` | 73 |
| `self.last_preview` | 44 |
| `self.show_file_channel` | 31 |
| `self.channel_dropdown` | 29 |
| `self.populate_thumbnails_for_channel` | 23 |
| `self.preview_canvas` | 19 |
| `self.collection_controller` | 18 |
| `self.filter_controller` | 16 |
| `self._spectros_loaded` | 15 |
| `self._schedule_marker_refresh` | 14 |
| `self.ensure_spectros_loaded` | 13 |
| `self.headers` | 13 |
| `self.frame_map_widget` | 10 |
| `self._invalidate_thumbnail_cache` | 10 |
| `self.show_collection_tray` | 10 |

## MultiPreviewCanvas

`sxm_viewer/gui/canvases/detail_preview_canvas.py`

- methods analysed: **525**
- distinct `self.` attributes: **717**
- cohesive groups (>= 3 attrs): **31**

### Candidate extractions, most isolated first

`Isolation` = share of methods touching this group that touch *nothing else*. High = safe to extract.

| Isolation | Attrs | Pure/Touching methods | Attributes |
|---|---|---|---|
| 80% | 16 | 12/15 | `_display_relative_zero_menu_state_callback`, `_virtual_copy_callback`, `_restore_windows_callback`, `_scale_bar_edit_view_key`, `_filter_menu_callback`, `_collection_menu_callback` +10 |
| 50% | 8 | 1/2 | `set_profile_label_bg_alpha`, `set_profile_show_all_labels`, `set_profile_label_scale`, `set_profile_label_placement`, `set_profile_label_mode`, `set_profile_ruler_visible` +2 |
| 50% | 4 | 3/6 | `_ch_height_style_callback`, `_ch_height_publication_callback`, `_ch_height_menu_callback`, `_ch_height_text_callback` |
| 43% | 3 | 3/7 | `_recent_svg_molecule_paths`, `_last_svg_molecule_dir_cb`, `_last_svg_molecule_dir` |
| 33% | 4 | 1/3 | `_resize_settle_timer`, `_resize_draft_timer`, `_resize_reflow_threshold_px`, `_last_resize_size` |
| 33% | 4 | 1/3 | `_scale_bar_manual_length`, `_format_scale_bar_label`, `_scale_bar_span`, `_calculate_best_scale_bar` |
| 29% | 14 | 2/7 | `_hover_view_ax`, `_view_font_scale`, `_active_view_ax`, `_colorbars`, `_draw_outlines`, `_async_redraw_once` +8 |
| 29% | 8 | 2/7 | `_outline_start`, `_outline_ax`, `_crop_rect`, `_crop_square`, `_crop_ax`, `_crop_start` +2 |
| 25% | 3 | 1/4 | `_apply_popup_style_callback`, `_apply_popup_style_tooltip`, `_apply_popup_style_label` |
| 22% | 5 | 2/9 | `_undo_history`, `_display_preset_base_state`, `_display_preset_transient`, `_active_display_preset`, `_undo_restore_in_progress` |
| 14% | 3 | 1/7 | `_profile_blit_active`, `_profile_background`, `_blit_profile_artists` |
| 7% | 7 | 1/15 | `_show_colorbar`, `_show_title`, `_show_ticks`, `_publication_mode`, `_fit_to_canvas`, `_frame_fill_mode` +1 |
| 4% | 4 | 1/24 | `_fixed_crop_template`, `_fixed_crop_template_bounds`, `_fixed_crop_template_view_key`, `_fixed_crop_template_pixel_bounds` |
| 0% | 14 | 0/7 | `_pan_active`, `_pan_start_lim`, `_molecule_gizmo_drag`, `_molecule_drag_start`, `_pan_ax`, `_molecule_drag_start_px` +8 |
| 0% | 9 | 0/5 | `set_show_profile_overlays`, `set_show_shortcut_hint`, `set_show_molecules`, `set_show_angle_overlays`, `enable_scale_bar`, `set_show_acquisition_overlay` +3 |
| 0% | 6 | 0/13 | `_profile_line`, `_profile_echo_artists`, `_profile_p0`, `_profile_p1`, `_profile_label`, `_profile_endpoint_labels` |
| 0% | 5 | 0/16 | `_active_profile_color`, `_active_profile_lw`, `_active_profile_line_style`, `_active_profile_marker_style`, `_active_profile_marker_size` |
| 0% | 5 | 0/7 | `_publication_colorbar_label`, `_apply_view_theme`, `_apply_view_font_scale`, `_add_scale_bar`, `_update_highlight_artists` |
| 0% | 4 | 0/16 | `_show_profile_overlays`, `_show_angle_overlays`, `_show_acquisition_overlay`, `_show_shortcut_hint` |
| 0% | 4 | 0/7 | `_dragging`, `_profile_marker_drag_idx`, `_saved_profile_drag`, `_line_drag_origin` |
| 0% | 4 | 0/6 | `_molecule_gizmo_axes`, `_molecule_gizmo_timer`, `_molecule_gizmo_until`, `_molecule_gizmo_artists` |
| 0% | 4 | 0/4 | `_outline_default_ls`, `_outline_default_lw`, `_outline_default_color`, `_outline_threshold` |
| 0% | 3 | 0/12 | `_profile_marker_key`, `_profile_marker_positions_by_key`, `_profile_marker_domain_by_key` |
| 0% | 3 | 0/10 | `_profile_label_bg_alpha`, `_profile_label_placement`, `_profile_ruler_visible` |
| 0% | 3 | 0/6 | `_build_profile_data`, `_profile_live_ref`, `_create_profile_id_label` |

### Group detail

<details><summary>16 attributes, 80% isolated (12/15 methods)</summary>

- `self._arrange_windows_callback` - 3 methods
- `self._close_windows_callback` - 3 methods
- `self._collection_help_callback` - 3 methods
- `self._collection_menu_callback` - 3 methods
- `self._display_relative_zero_menu_callback` - 3 methods
- `self._display_relative_zero_menu_state_callback` - 3 methods
- `self._display_relative_zero_menu_tooltip` - 3 methods
- `self._filter_menu_callback` - 3 methods
- `self._histogram_dialog_callback` - 3 methods
- `self._histogram_reset_callback` - 3 methods
- `self._minimize_windows_callback` - 3 methods
- `self._relative_axes_menu_callback` - 3 methods
- `self._restore_windows_callback` - 3 methods
- `self._scale_bar_edit_view_key` - 3 methods
- `self._spectra_compare_all_callback` - 3 methods
- `self._virtual_copy_callback` - 3 methods

</details>

<details><summary>8 attributes, 50% isolated (1/2 methods)</summary>

- `self.set_profile_label_bg_alpha` - 2 methods
- `self.set_profile_label_mode` - 2 methods
- `self.set_profile_label_orientation` - 2 methods
- `self.set_profile_label_placement` - 2 methods
- `self.set_profile_label_scale` - 2 methods
- `self.set_profile_overlay_outline` - 2 methods
- `self.set_profile_ruler_visible` - 2 methods
- `self.set_profile_show_all_labels` - 2 methods

</details>

<details><summary>4 attributes, 50% isolated (3/6 methods)</summary>

- `self._ch_height_menu_callback` - 3 methods
- `self._ch_height_publication_callback` - 3 methods
- `self._ch_height_style_callback` - 3 methods
- `self._ch_height_text_callback` - 3 methods

</details>

<details><summary>3 attributes, 43% isolated (3/7 methods)</summary>

- `self._last_svg_molecule_dir` - 4 methods
- `self._last_svg_molecule_dir_cb` - 4 methods
- `self._recent_svg_molecule_paths` - 5 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self._last_resize_size` - 2 methods
- `self._resize_draft_timer` - 3 methods
- `self._resize_reflow_threshold_px` - 2 methods
- `self._resize_settle_timer` - 3 methods

</details>

<details><summary>4 attributes, 33% isolated (1/3 methods)</summary>

- `self._calculate_best_scale_bar` - 2 methods
- `self._format_scale_bar_label` - 2 methods
- `self._scale_bar_manual_length` - 3 methods
- `self._scale_bar_span` - 2 methods

</details>

<details><summary>14 attributes, 29% isolated (2/7 methods)</summary>

- `self._active_view_ax` - 4 methods
- `self._async_redraw_once` - 2 methods
- `self._colorbars` - 4 methods
- `self._compose_view_title` - 2 methods
- `self._draw_acquisition_overlay` - 2 methods
- `self._draw_filter_summary_overlay` - 2 methods
- `self._draw_fixed_crop_history` - 2 methods
- `self._draw_outlines` - 2 methods
- `self._draw_shortcut_hint` - 2 methods
- `self._draw_spectra` - 2 methods
- `self._hover_view_ax` - 4 methods
- `self._style_publication_colorbar` - 2 methods
- `self._suppress_internal_draw_requests` - 2 methods
- `self._view_font_scale` - 4 methods

</details>

<details><summary>8 attributes, 29% isolated (2/7 methods)</summary>

- `self._crop_ax` - 5 methods
- `self._crop_last_ts` - 3 methods
- `self._crop_rect` - 5 methods
- `self._crop_square` - 5 methods
- `self._crop_start` - 5 methods
- `self._outline_ax` - 6 methods
- `self._outline_rect` - 4 methods
- `self._outline_start` - 6 methods

</details>

<details><summary>3 attributes, 25% isolated (1/4 methods)</summary>

- `self._apply_popup_style_callback` - 4 methods
- `self._apply_popup_style_label` - 3 methods
- `self._apply_popup_style_tooltip` - 3 methods

</details>

<details><summary>5 attributes, 22% isolated (2/9 methods)</summary>

- `self._active_display_preset` - 6 methods
- `self._display_preset_base_state` - 6 methods
- `self._display_preset_transient` - 6 methods
- `self._undo_history` - 6 methods
- `self._undo_restore_in_progress` - 3 methods

</details>

<details><summary>3 attributes, 14% isolated (1/7 methods)</summary>

- `self._blit_profile_artists` - 3 methods
- `self._profile_background` - 6 methods
- `self._profile_blit_active` - 6 methods

</details>

<details><summary>7 attributes, 7% isolated (1/15 methods)</summary>

- `self._fit_to_canvas` - 7 methods
- `self._frame_fill_mode` - 6 methods
- `self._frame_fill_prev_state` - 5 methods
- `self._publication_mode` - 8 methods
- `self._show_colorbar` - 9 methods
- `self._show_ticks` - 9 methods
- `self._show_title` - 9 methods

</details>

### Extraction manifests (most isolated groups)

`pure` methods touch only this group's attributes and can move wholesale. `mixed` methods touch the group *and* other state - they stay put and call into the extracted object.

<details><summary><b>16 attrs, 80% isolated</b> - `_arrange_windows_callback`, `_close_windows_callback`, `_collection_help_callback`, `_collection_menu_callback` ...</summary>

**State to move (16):**
- `self._arrange_windows_callback`
- `self._close_windows_callback`
- `self._collection_help_callback`
- `self._collection_menu_callback`
- `self._display_relative_zero_menu_callback`
- `self._display_relative_zero_menu_state_callback`
- `self._display_relative_zero_menu_tooltip`
- `self._filter_menu_callback`
- `self._histogram_dialog_callback`
- `self._histogram_reset_callback`
- `self._minimize_windows_callback`
- `self._relative_axes_menu_callback`
- `self._restore_windows_callback`
- `self._scale_bar_edit_view_key`
- `self._spectra_compare_all_callback`
- `self._virtual_copy_callback`

**Pure methods (12) - move these:**
- `set_collection_menu_callback()`
- `set_display_relative_zero_menu_callback()`
- `set_filter_menu_callback()`
- `set_histogram_dialog_callback()`
- `set_histogram_reset_callback()`
- `set_relative_axes_menu_callback()`
- `set_spectra_compare_all_callback()`
- `set_virtual_copy_callback()`
- `set_window_arrange_callback()`
- `set_window_close_callback()`
- `set_window_minimize_callback()`
- `set_window_restore_callback()`

**Mixed methods (3) - keep, delegate:**
- `__init__()` - also touches `_active_angle_frame_idx`, `_active_display_preset`, `_active_molecule_idx`, `_active_profile_color`, `_active_profile_line_style` ...
- `_show_context_menu()` - also touches `_add_context_cmap_menu`, `_add_context_compare_menu`, `_add_context_crop_menu`, `_add_context_measure_menu`, `_add_context_molecules_menu` ...
- `_show_sb_context_menu()` - also touches `_ax_view_map`, `_pick_sb_bar_color`, `_pick_sb_text_color`, `_populate_sb_length_menu`, `_scale_bar_font_weight` ...

</details>

<details><summary><b>8 attrs, 50% isolated</b> - `set_profile_label_bg_alpha`, `set_profile_label_mode`, `set_profile_label_orientation`, `set_profile_label_placement` ...</summary>

**State to move (8):**
- `self.set_profile_label_bg_alpha`
- `self.set_profile_label_mode`
- `self.set_profile_label_orientation`
- `self.set_profile_label_placement`
- `self.set_profile_label_scale`
- `self.set_profile_overlay_outline`
- `self.set_profile_ruler_visible`
- `self.set_profile_show_all_labels`

**Pure methods (1) - move these:**
- `_restore_profile_display_favorite()`

**Mixed methods (1) - keep, delegate:**
- `_add_context_profile_labels_menu()` - also touches `_profile_label_bg_alpha`, `_profile_label_mode`, `_profile_label_orientation`, `_profile_label_placement`, `_profile_label_scale` ...

</details>

<details><summary><b>4 attrs, 50% isolated</b> - `_ch_height_menu_callback`, `_ch_height_publication_callback`, `_ch_height_style_callback`, `_ch_height_text_callback`</summary>

**State to move (4):**
- `self._ch_height_menu_callback`
- `self._ch_height_publication_callback`
- `self._ch_height_style_callback`
- `self._ch_height_text_callback`

**Pure methods (3) - move these:**
- `_ch_height_overlay_text()`
- `_ch_label_style()`
- `set_ch_height_label_callbacks()`

**Mixed methods (3) - keep, delegate:**
- `__init__()` - also touches `_active_angle_frame_idx`, `_active_display_preset`, `_active_molecule_idx`, `_active_profile_color`, `_active_profile_line_style` ...
- `_ch_height_in_publication()` - also touches `_publication_mode`
- `_show_context_menu()` - also touches `_add_context_cmap_menu`, `_add_context_compare_menu`, `_add_context_crop_menu`, `_add_context_measure_menu`, `_add_context_molecules_menu` ...

</details>

<details><summary><b>3 attrs, 43% isolated</b> - `_last_svg_molecule_dir`, `_last_svg_molecule_dir_cb`, `_recent_svg_molecule_paths`</summary>

**State to move (3):**
- `self._last_svg_molecule_dir`
- `self._last_svg_molecule_dir_cb`
- `self._recent_svg_molecule_paths`

**Pure methods (3) - move these:**
- `get_recent_svg_molecule_paths()`
- `set_last_svg_molecule_dir()`
- `set_last_svg_molecule_dir_callback()`

**Mixed methods (4) - keep, delegate:**
- `__init__()` - also touches `_active_angle_frame_idx`, `_active_display_preset`, `_active_molecule_idx`, `_active_profile_color`, `_active_profile_line_style` ...
- `_add_context_molecules_menu()` - also touches `_clear_molecules`, `_clear_svg_molecules`, `_load_molecule_dialog`, `_load_svg_molecule_dialog`, `_recent_molecule_paths` ...
- `_load_svg_molecule_dialog()` - also touches `add_svg_molecule`
- `add_svg_molecule()` - also touches `_active_svg_molecule_idx`, `_recent_svg_molecule_cb`, `_redraw`, `_svg_overlay_angstrom_per_unit`, `_svg_overlay_reference_view` ...

</details>

<details><summary><b>4 attrs, 33% isolated</b> - `_last_resize_size`, `_resize_draft_timer`, `_resize_reflow_threshold_px`, `_resize_settle_timer`</summary>

**State to move (4):**
- `self._last_resize_size`
- `self._resize_draft_timer`
- `self._resize_reflow_threshold_px`
- `self._resize_settle_timer`

**Pure methods (1) - move these:**
- `resizeEvent()`

**Mixed methods (2) - keep, delegate:**
- `__init__()` - also touches `_active_angle_frame_idx`, `_active_display_preset`, `_active_molecule_idx`, `_active_profile_color`, `_active_profile_line_style` ...
- `_finalize_after_resize()` - also touches `_redraw`

</details>

<details><summary><b>4 attrs, 33% isolated</b> - `_calculate_best_scale_bar`, `_format_scale_bar_label`, `_scale_bar_manual_length`, `_scale_bar_span`</summary>

**State to move (4):**
- `self._calculate_best_scale_bar`
- `self._format_scale_bar_label`
- `self._scale_bar_manual_length`
- `self._scale_bar_span`

**Pure methods (1) - move these:**
- `_calculate_best_scale_bar()`

**Mixed methods (2) - keep, delegate:**
- `_add_scale_bar()` - also touches `_plot_style_state`, `_scale_bar_artists`, `_scale_bar_auto_color`, `_scale_bar_font_size`, `_scale_bar_font_weight` ...
- `_populate_sb_length_menu()` - also touches `_nice_scale_bar_lengths`, `_prompt_sb_length`, `_set_sb_length`

</details>

<details><summary><b>14 attrs, 29% isolated</b> - `_active_view_ax`, `_async_redraw_once`, `_colorbars`, `_compose_view_title` ...</summary>

**State to move (14):**
- `self._active_view_ax`
- `self._async_redraw_once`
- `self._colorbars`
- `self._compose_view_title`
- `self._draw_acquisition_overlay`
- `self._draw_filter_summary_overlay`
- `self._draw_fixed_crop_history`
- `self._draw_outlines`
- `self._draw_shortcut_hint`
- `self._draw_spectra`
- `self._hover_view_ax`
- `self._style_publication_colorbar`
- `self._suppress_internal_draw_requests`
- `self._view_font_scale`

**Pure methods (2) - move these:**
- `_set_active_view_ax()`
- `_set_hover_view_ax()`

**Mixed methods (5) - keep, delegate:**
- `__init__()` - also touches `_active_angle_frame_idx`, `_active_display_preset`, `_active_molecule_idx`, `_active_profile_color`, `_active_profile_line_style` ...
- `_apply_view_theme()` - also touches `_compute_theme_sig`, `_detail_dark`, `_detail_grid`, `_recolor_scale_bars`, `_theme_sig` ...
- `_fast_update_single_view()` - also touches `_add_scale_bar`, `_apply_view_font_scale`, `_apply_view_theme`, `_ax_view_map`, `_compute_font_sig` ...
- `_redraw()` - also touches `_add_scale_bar`, `_angle_frames`, `_apply_angle_visibility`, `_apply_profile_visibility`, `_apply_tight_layout_safe` ...
- `wheelEvent()` - also touches `_apply_view_font_scale`

</details>

<details><summary><b>8 attrs, 29% isolated</b> - `_crop_ax`, `_crop_last_ts`, `_crop_rect`, `_crop_square` ...</summary>

**State to move (8):**
- `self._crop_ax`
- `self._crop_last_ts`
- `self._crop_rect`
- `self._crop_square`
- `self._crop_start`
- `self._outline_ax`
- `self._outline_rect`
- `self._outline_start`

**Pure methods (2) - move these:**
- `_reset_crop_state()`
- `_reset_outline_state()`

**Mixed methods (5) - keep, delegate:**
- `__init__()` - also touches `_active_angle_frame_idx`, `_active_display_preset`, `_active_molecule_idx`, `_active_profile_color`, `_active_profile_line_style` ...
- `_finish_outline_drag()` - also touches `_ax_view_map`, `_outline_from_point`, `_reset_outline_state`
- `_on_base_click()` - also touches `_add_angle_frame_at`, `_apply_fixed_crop_quick`, `_ax_view_map`, `_begin_fixed_crop_template_drag`, `_begin_molecule_gizmo_drag` ...
- `_on_crop_release()` - also touches `_ax_view_map`, `_axis_coord_to_pixel`, `_compute_crop_extent`, `_crop_callback`, `_display_extent_for_view` ...
- `_on_motion_value()` - also touches `_ax_view_map`, `_event_qt_modifiers`, `_fixed_crop_drag_last_ts`, `_fixed_crop_drag_throttle_ms`, `_fixed_crop_template_drag` ...

</details>

### Spine attributes (touched by the most methods)

> These resist extraction and define what the class is *actually* about; everything else is a passenger.

| Attribute | Methods touching |
|---|---|
| `self.draw_idle` | 67 |
| `self.main_ax` | 58 |
| `self._redraw` | 48 |
| `self.push_undo_state` | 44 |
| `self._ax_view_map` | 34 |
| `self._saved_profiles` | 31 |
| `self.profile_pts` | 26 |
| `self._notify_views_callback` | 25 |
| `self.views` | 23 |
| `self._fixed_crop_template` | 22 |
| `self._angle_frames` | 20 |
| `self.molecules` | 20 |
| `self._outline_key` | 20 |
| `self.fig` | 17 |
| `self._fixed_crop_template_bounds` | 17 |

