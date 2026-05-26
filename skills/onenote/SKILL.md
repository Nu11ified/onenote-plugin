---
name: onenote
description: Use Microsoft OneNote from Codex for structured note capture, retrieval, page creation, and page updates through Microsoft Graph.
---

# OneNote Skill

Use this skill when the user asks Codex to read, search, organize, create, or update OneNote notebooks, sections, pages, meeting notes, research notes, task lists, project logs, or other structured notes.

## Authentication

The bundled MCP server uses Microsoft Graph delegated authentication. App-only OneNote access is not supported for new integrations.

Before using notebook tools:

1. Check auth with `onenote_auth_status`.
2. If not signed in, call `onenote_auth_start`.
3. Tell the user to open the returned verification URL and enter the returned user code.
4. Call `onenote_auth_complete` after the user signs in.

The OAuth app must be a Microsoft Entra public client with delegated Graph permissions for `Notes.ReadWrite`, `offline_access`, and `User.Read`. The client ID can be supplied with the `ONENOTE_CLIENT_ID` environment variable or as the `clientId` argument to `onenote_auth_start`.

## Structured Workflows

Prefer OneNote pages that are easy to scan and update:

- Use one page per meeting, decision log, project update, research brief, or task plan.
- Use a clear title, short summary, dated context, sections, and action items.
- Use semantic HTML fragments with headings, paragraphs, tables, and lists.
- Add stable `data-id` attributes to important blocks when future updates are likely.

For new pages, use `onenote_create_page` with either a target `sectionId` or a `sectionName`. When the user only names a notebook or section, list notebooks and sections first, then ask for clarification only if multiple plausible matches remain.

For existing pages, fetch content with `onenote_get_page_content` and `includeIDs: true` before using generated element IDs in replace operations. Prefer `data-id` targets for append and insert updates when available.

## Tool Selection

- `onenote_list_notebooks`: discover available notebooks.
- `onenote_list_sections`: discover sections globally or within a notebook.
- `onenote_list_pages`: inspect recent pages globally or within a section.
- `onenote_get_page_content`: read full page HTML.
- `onenote_create_page`: create a OneNote page from HTML.
- `onenote_append_to_page`: append a section, note, or action item to an existing page.
- `onenote_update_page`: send advanced OneNote PATCH change objects.

Keep page HTML concise. OneNote supports a subset of HTML and CSS, so avoid scripts, complex layouts, external styling, and unsupported interactive elements.
