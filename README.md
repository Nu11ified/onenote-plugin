# OneNote Codex Plugin

This plugin gives Codex delegated Microsoft Graph tools for OneNote:

- Start and complete device-code sign-in.
- List notebooks, sections, and recent pages.
- Fetch OneNote page HTML.
- Create structured pages from HTML.
- Append or patch existing pages.

## Setup

Create a Microsoft Entra app registration configured as a public client, then grant delegated Microsoft Graph permissions:

- `Notes.ReadWrite`
- `offline_access`
- `User.Read`

Set the client ID before enabling the bundled MCP server:

```sh
export ONENOTE_CLIENT_ID="00000000-0000-0000-0000-000000000000"
```

Optional environment variables:

```sh
export ONENOTE_TENANT_ID="common"
export ONENOTE_SCOPES="Notes.ReadWrite offline_access User.Read"
```

Microsoft Graph OneNote app-only authentication is not supported for this plugin because Microsoft ended app-only support for the OneNote API on March 31, 2025.

## Local Validation

```sh
node scripts/onenote-mcp.js --self-test
python3 /Users/saimanasr/.codex/skills/.system/plugin-creator/scripts/validate_plugin.py /Users/saimanasr/Documents/onenote-plugin
```
