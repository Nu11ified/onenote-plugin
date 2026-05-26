# OneNote Codex Plugin

This plugin gives Codex delegated Microsoft Graph tools for OneNote:

- Open Microsoft browser sign-in with OAuth PKCE.
- Store refresh tokens in macOS Keychain.
- Start and complete device-code sign-in as a fallback.
- List notebooks, sections, and recent pages.
- Fetch OneNote page HTML.
- Create structured pages from HTML.
- Append or patch existing pages.

## Setup

Create a Microsoft Entra app registration configured as a public client, then grant delegated Microsoft Graph permissions:

- `Notes.ReadWrite`
- `offline_access`
- `User.Read`

Add `http://localhost` as a redirect URI for the public client flow. The plugin starts a temporary local callback server on a dynamic port and opens Microsoft sign-in in your browser.

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

## Sign In

Use `onenote_login` from Codex. It opens your browser, waits for the localhost callback, stores the refresh token in macOS Keychain, and keeps non-sensitive token metadata under the plugin data directory.

If browser sign-in is unavailable, use the fallback device-code tools:

1. `onenote_auth_start`
2. Open the returned verification URL and enter the code.
3. `onenote_auth_complete`

## Local Validation

```sh
node scripts/onenote-mcp.js --self-test
python3 /Users/saimanasr/.codex/skills/.system/plugin-creator/scripts/validate_plugin.py /Users/saimanasr/Documents/onenote-plugin
```
