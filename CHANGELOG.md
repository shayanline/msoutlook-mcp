# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.3.1] - 2026-06-04

### Changed
- Documentation: point the `msteams-mcp` reference at `@shayanline/msteams-mcp`.

## [0.3.0] - 2026-06-04

### Added
- Review-first drafts: `outlook_create_reply_draft`, `outlook_create_forward_draft`, `outlook_update_draft`.
- Attachments: send files with `outlook_send_email`/`outlook_create_draft`, plus `outlook_add_attachment`, `outlook_list_attachments`, `outlook_save_attachment`.
- Conversation, categories and folders: `outlook_get_conversation`, `outlook_set_categories`, `outlook_create_folder`, `outlook_rename_folder`, `outlook_delete_folder`.
- Calendar scheduling: `outlook_get_schedule`, `outlook_find_meeting_times`, `outlook_cancel_event`, `outlook_forward_event`.
- People and org directory: `outlook_get_availability`, `outlook_get_user_profile`, `outlook_get_manager`, `outlook_get_direct_reports`, `outlook_get_user_photo`.
- Out of office: `outlook_get_automatic_replies`, `outlook_set_automatic_replies`.

### Changed
- Email bodies default to HTML and plain-text newlines auto-convert to `<br>` so messages no longer collapse into one block.

## [0.1.x]

See the [GitHub releases](https://github.com/shayanline/msoutlook-mcp/releases) for earlier history.

[Unreleased]: https://github.com/shayanline/msoutlook-mcp/compare/v0.3.1...HEAD
[0.3.1]: https://github.com/shayanline/msoutlook-mcp/compare/v0.3.0...v0.3.1
[0.3.0]: https://github.com/shayanline/msoutlook-mcp/compare/v0.1.3...v0.3.0
