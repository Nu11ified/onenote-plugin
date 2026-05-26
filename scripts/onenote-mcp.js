#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DEFAULT_TENANT = process.env.ONENOTE_TENANT_ID || "common";
const DEFAULT_SCOPES = process.env.ONENOTE_SCOPES || "Notes.ReadWrite offline_access User.Read";
const DATA_DIR =
  process.env.PLUGIN_DATA ||
  process.env.ONENOTE_DATA_DIR ||
  path.join(os.homedir(), ".codex-onenote-plugin");
const TOKEN_PATH = path.join(DATA_DIR, "token.json");
const PENDING_AUTH_PATH = path.join(DATA_DIR, "pending-auth.json");

const tools = [
  {
    name: "onenote_auth_status",
    description: "Show whether OneNote delegated Microsoft Graph auth is configured and cached.",
    inputSchema: objectSchema({})
  },
  {
    name: "onenote_auth_start",
    description: "Start Microsoft device-code sign-in for OneNote delegated Graph access.",
    inputSchema: objectSchema({
      clientId: stringProp("Microsoft Entra application client ID. Defaults to ONENOTE_CLIENT_ID."),
      tenantId: stringProp("Tenant ID or common/organizations/consumers. Defaults to ONENOTE_TENANT_ID or common."),
      scopes: stringProp("Space-delimited delegated Graph scopes. Defaults to Notes.ReadWrite offline_access User.Read.")
    })
  },
  {
    name: "onenote_auth_complete",
    description: "Complete a pending Microsoft device-code sign-in and cache the delegated token.",
    inputSchema: objectSchema({
      waitSeconds: numberProp("Maximum seconds to poll for completion. Defaults to 90.", 1, 900)
    })
  },
  {
    name: "onenote_clear_auth",
    description: "Delete cached OneNote auth tokens and pending device-code state.",
    inputSchema: objectSchema({})
  },
  {
    name: "onenote_list_notebooks",
    description: "List OneNote notebooks available to the signed-in user.",
    inputSchema: objectSchema({
      limit: numberProp("Maximum notebooks to return. Defaults to 50.", 1, 100)
    })
  },
  {
    name: "onenote_list_sections",
    description: "List OneNote sections globally or within a specific notebook.",
    inputSchema: objectSchema({
      notebookId: stringProp("Optional notebook ID."),
      limit: numberProp("Maximum sections to return. Defaults to 100.", 1, 100)
    })
  },
  {
    name: "onenote_list_pages",
    description: "List recent OneNote pages globally or within a specific section.",
    inputSchema: objectSchema({
      sectionId: stringProp("Optional section ID."),
      query: stringProp("Optional case-insensitive title filter applied client-side."),
      limit: numberProp("Maximum pages to return. Defaults to 25.", 1, 100)
    })
  },
  {
    name: "onenote_get_page_content",
    description: "Get a OneNote page HTML body, optionally including generated update IDs.",
    inputSchema: objectSchema({
      pageId: requiredStringProp("OneNote page ID."),
      includeIDs: booleanProp("Include generated element IDs for PATCH updates. Defaults to false."),
      maxChars: numberProp("Maximum HTML characters to return. Defaults to 30000.", 1000, 200000)
    }, ["pageId"])
  },
  {
    name: "onenote_create_page",
    description: "Create a OneNote page from a title and HTML fragment or full XHTML document.",
    inputSchema: objectSchema({
      title: requiredStringProp("Page title."),
      htmlBody: requiredStringProp("HTML fragment for the body, or a full XHTML document."),
      sectionId: stringProp("Target section ID. Preferred when known."),
      sectionName: stringProp("Top-level default-notebook section name. Created by Graph if missing."),
      createdIso: stringProp("Optional ISO timestamp for the OneNote created meta tag.")
    }, ["title", "htmlBody"])
  },
  {
    name: "onenote_append_to_page",
    description: "Append or prepend an HTML fragment to a OneNote page target.",
    inputSchema: objectSchema({
      pageId: requiredStringProp("OneNote page ID."),
      htmlFragment: requiredStringProp("Well-formed HTML fragment to append."),
      target: stringProp("Target element: body, #data-id, or generated ID. Defaults to body."),
      position: stringProp("append position: after for last child, before for first child. Defaults to after.")
    }, ["pageId", "htmlFragment"])
  },
  {
    name: "onenote_update_page",
    description: "Apply advanced OneNote PATCH change objects to a page.",
    inputSchema: objectSchema({
      pageId: requiredStringProp("OneNote page ID."),
      changes: {
        type: "array",
        description: "OneNote PATCH change objects with target, action, content, and optional position.",
        items: { type: "object" }
      }
    }, ["pageId", "changes"])
  }
];

function objectSchema(properties, required = []) {
  return {
    type: "object",
    properties,
    required,
    additionalProperties: false
  };
}

function stringProp(description) {
  return { type: "string", description };
}

function requiredStringProp(description) {
  return { type: "string", description, minLength: 1 };
}

function numberProp(description, minimum, maximum) {
  return { type: "number", description, minimum, maximum };
}

function booleanProp(description) {
  return { type: "boolean", description };
}

function ensureDataDir() {
  fs.mkdirSync(DATA_DIR, { recursive: true, mode: 0o700 });
}

function readJson(file) {
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") return null;
    throw error;
  }
}

function writeJson(file, value) {
  ensureDataDir();
  fs.writeFileSync(file, JSON.stringify(value, null, 2), { mode: 0o600 });
}

function unlinkIfExists(file) {
  try {
    fs.unlinkSync(file);
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
}

function authConfig(args = {}) {
  return {
    clientId: args.clientId || process.env.ONENOTE_CLIENT_ID,
    tenantId: args.tenantId || DEFAULT_TENANT,
    scopes: args.scopes || DEFAULT_SCOPES
  };
}

function tokenEndpoint(tenantId) {
  return `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
}

async function postForm(url, values) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(values)
  });
  const text = await response.text();
  const payload = text ? JSON.parse(text) : {};
  if (!response.ok) {
    const message = payload.error_description || payload.error || text || response.statusText;
    const error = new Error(message);
    error.status = response.status;
    error.payload = payload;
    throw error;
  }
  return payload;
}

async function startAuth(args) {
  const config = authConfig(args);
  if (!config.clientId) {
    throw new Error(
      "Missing Microsoft client ID. Set ONENOTE_CLIENT_ID or pass clientId. The app must allow public client/device-code auth."
    );
  }

  const payload = await postForm(
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/devicecode`,
    {
      client_id: config.clientId,
      scope: config.scopes
    }
  );

  writeJson(PENDING_AUTH_PATH, {
    ...config,
    deviceCode: payload.device_code,
    interval: payload.interval || 5,
    expiresAt: Date.now() + payload.expires_in * 1000
  });

  return {
    status: "pending",
    verificationUri: payload.verification_uri,
    userCode: payload.user_code,
    expiresInSeconds: payload.expires_in,
    message: payload.message
  };
}

async function completeAuth(args) {
  const pending = readJson(PENDING_AUTH_PATH);
  if (!pending) {
    throw new Error("No pending OneNote auth flow. Run onenote_auth_start first.");
  }
  if (Date.now() > pending.expiresAt) {
    unlinkIfExists(PENDING_AUTH_PATH);
    throw new Error("The pending OneNote auth code expired. Run onenote_auth_start again.");
  }

  const deadline = Date.now() + Math.min(args.waitSeconds || 90, 900) * 1000;
  let intervalMs = Math.max(pending.interval || 5, 1) * 1000;

  while (Date.now() <= deadline) {
    try {
      const token = await postForm(tokenEndpoint(pending.tenantId), {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: pending.clientId,
        device_code: pending.deviceCode
      });
      writeToken(token, pending);
      unlinkIfExists(PENDING_AUTH_PATH);
      return {
        status: "signed_in",
        scopes: token.scope,
        expiresAt: new Date(Date.now() + token.expires_in * 1000).toISOString()
      };
    } catch (error) {
      const code = error.payload && error.payload.error;
      if (code === "authorization_pending") {
        await sleep(intervalMs);
        continue;
      }
      if (code === "slow_down") {
        intervalMs += 5000;
        await sleep(intervalMs);
        continue;
      }
      if (code === "authorization_declined" || code === "expired_token") {
        unlinkIfExists(PENDING_AUTH_PATH);
      }
      throw error;
    }
  }

  return {
    status: "pending",
    message: "Sign-in has not completed yet. Call onenote_auth_complete again after the browser sign-in finishes."
  };
}

function writeToken(token, config) {
  writeJson(TOKEN_PATH, {
    accessToken: token.access_token,
    refreshToken: token.refresh_token,
    tokenType: token.token_type,
    scopes: token.scope,
    clientId: config.clientId,
    tenantId: config.tenantId,
    expiresAt: Date.now() + token.expires_in * 1000
  });
}

async function accessToken() {
  const token = readJson(TOKEN_PATH);
  if (!token) {
    throw new Error("OneNote is not signed in. Run onenote_auth_start and onenote_auth_complete first.");
  }
  if (token.expiresAt && Date.now() < token.expiresAt - 120000) {
    return token.accessToken;
  }
  if (!token.refreshToken || !token.clientId) {
    throw new Error("Cached OneNote token cannot be refreshed. Run onenote_auth_start again.");
  }

  const refreshed = await postForm(tokenEndpoint(token.tenantId || DEFAULT_TENANT), {
    grant_type: "refresh_token",
    client_id: token.clientId,
    refresh_token: token.refreshToken,
    scope: token.scopes || DEFAULT_SCOPES
  });
  writeToken(refreshed, token);
  return refreshed.access_token;
}

function authStatus() {
  const token = readJson(TOKEN_PATH);
  const pending = readJson(PENDING_AUTH_PATH);
  return {
    signedIn: Boolean(token && token.accessToken),
    expiresAt: token && token.expiresAt ? new Date(token.expiresAt).toISOString() : null,
    scopes: token ? token.scopes : null,
    pendingAuth: Boolean(pending),
    hasClientId: Boolean(process.env.ONENOTE_CLIENT_ID || (pending && pending.clientId) || (token && token.clientId)),
    dataDir: DATA_DIR
  };
}

async function graphRequest(method, resource, options = {}) {
  const url = resource.startsWith("https://") ? resource : `${GRAPH_ROOT}${resource}`;
  const headers = {
    Authorization: `Bearer ${await accessToken()}`,
    Accept: options.accept || "application/json",
    ...options.headers
  };
  const response = await fetch(url, {
    method,
    headers,
    body: options.body
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Graph ${method} ${url} failed with ${response.status}: ${truncate(text, 2000)}`);
  }
  const contentType = response.headers.get("content-type") || "";
  if (options.rawText || contentType.includes("text/html") || contentType.includes("application/xhtml")) {
    return text;
  }
  return text ? JSON.parse(text) : null;
}

async function listCollection(resource, limit) {
  const separator = resource.includes("?") ? "&" : "?";
  const payload = await graphRequest("GET", `${resource}${separator}$top=${encodeURIComponent(String(limit))}`);
  return {
    value: payload.value || [],
    nextLink: payload["@odata.nextLink"] || null
  };
}

async function callTool(name, args = {}) {
  switch (name) {
    case "onenote_auth_status":
      return authStatus();
    case "onenote_auth_start":
      return startAuth(args);
    case "onenote_auth_complete":
      return completeAuth(args);
    case "onenote_clear_auth":
      unlinkIfExists(TOKEN_PATH);
      unlinkIfExists(PENDING_AUTH_PATH);
      return { cleared: true };
    case "onenote_list_notebooks":
      return listCollection("/me/onenote/notebooks", args.limit || 50);
    case "onenote_list_sections": {
      const base = args.notebookId
        ? `/me/onenote/notebooks/${encodeURIComponent(args.notebookId)}/sections`
        : "/me/onenote/sections";
      return listCollection(base, args.limit || 100);
    }
    case "onenote_list_pages": {
      const base = args.sectionId
        ? `/me/onenote/sections/${encodeURIComponent(args.sectionId)}/pages?pagelevel=true`
        : "/me/onenote/pages?pagelevel=true";
      const result = await listCollection(base, args.limit || 25);
      if (args.query) {
        const query = args.query.toLowerCase();
        result.value = result.value.filter((page) => (page.title || "").toLowerCase().includes(query));
      }
      return result;
    }
    case "onenote_get_page_content": {
      const include = args.includeIDs ? "?includeIDs=true" : "";
      const html = await graphRequest(
        "GET",
        `/me/onenote/pages/${encodeURIComponent(args.pageId)}/content${include}`,
        { accept: "text/html", rawText: true }
      );
      return {
        pageId: args.pageId,
        html: truncate(html, args.maxChars || 30000),
        truncated: html.length > (args.maxChars || 30000)
      };
    }
    case "onenote_create_page": {
      const resource = createPageResource(args);
      const html = pageDocument(args.title, args.htmlBody, args.createdIso);
      return graphRequest("POST", resource, {
        headers: { "Content-Type": "application/xhtml+xml" },
        body: html
      });
    }
    case "onenote_append_to_page": {
      const changes = [
        {
          target: args.target || "body",
          action: "append",
          position: args.position || "after",
          content: args.htmlFragment
        }
      ];
      await patchPage(args.pageId, changes);
      return { updated: true, pageId: args.pageId, changesApplied: 1 };
    }
    case "onenote_update_page":
      if (!Array.isArray(args.changes)) {
        throw new Error("changes must be an array of OneNote PATCH change objects.");
      }
      await patchPage(args.pageId, args.changes);
      return { updated: true, pageId: args.pageId, changesApplied: args.changes.length };
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

function createPageResource(args) {
  if (args.sectionId) {
    return `/me/onenote/sections/${encodeURIComponent(args.sectionId)}/pages`;
  }
  if (args.sectionName) {
    return `/me/onenote/pages?sectionName=${encodeURIComponent(args.sectionName)}`;
  }
  return "/me/onenote/pages";
}

async function patchPage(pageId, changes) {
  return graphRequest("PATCH", `/me/onenote/pages/${encodeURIComponent(pageId)}/content`, {
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(changes)
  });
}

function pageDocument(title, htmlBody, createdIso) {
  if (/^\s*(<!doctype|<html[\s>])/i.test(htmlBody)) {
    return htmlBody;
  }
  const created = createdIso ? `<meta name="created" content="${escapeAttr(createdIso)}" />` : "";
  return [
    "<!DOCTYPE html>",
    '<html xmlns="http://www.w3.org/1999/xhtml">',
    "<head>",
    `<title>${escapeHtml(title)}</title>`,
    created,
    "</head>",
    "<body>",
    htmlBody,
    "</body>",
    "</html>"
  ].join("\n");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/"/g, "&quot;");
}

function truncate(value, maxChars) {
  const text = String(value);
  return text.length > maxChars ? text.slice(0, maxChars) : text;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function send(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function result(id, value) {
  send({ jsonrpc: "2.0", id, result: value });
}

function errorResult(id, code, message) {
  send({ jsonrpc: "2.0", id, error: { code, message } });
}

async function handle(message) {
  if (!message || message.jsonrpc !== "2.0") return;
  const { id, method, params = {} } = message;

  try {
    if (method === "notifications/initialized") return;
    if (method === "initialize") {
      return result(id, {
        protocolVersion: params.protocolVersion || "2024-11-05",
        capabilities: { tools: {} },
        serverInfo: { name: "onenote", version: "0.1.0" }
      });
    }
    if (method === "ping") {
      return result(id, {});
    }
    if (method === "tools/list") {
      return result(id, { tools });
    }
    if (method === "tools/call") {
      const payload = await callTool(params.name, params.arguments || {});
      return result(id, {
        content: [{ type: "text", text: JSON.stringify(payload, null, 2) }]
      });
    }
    if (id !== undefined) errorResult(id, -32601, `Method not found: ${method}`);
  } catch (error) {
    if (id !== undefined) errorResult(id, -32000, error.message || String(error));
  }
}

function startServer() {
  let buffer = "";
  process.stdin.setEncoding("utf8");
  process.stdin.on("data", (chunk) => {
    buffer += chunk;
    let newline;
    while ((newline = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, newline).replace(/\r$/, "");
      buffer = buffer.slice(newline + 1);
      if (!line.trim()) continue;
      try {
        handle(JSON.parse(line));
      } catch (error) {
        errorResult(null, -32700, error.message || String(error));
      }
    }
  });
}

if (process.argv.includes("--self-test")) {
  if (tools.length < 10) throw new Error("Expected OneNote MCP tools to be registered.");
  for (const tool of tools) {
    if (!tool.name || !tool.inputSchema || tool.inputSchema.type !== "object") {
      throw new Error(`Invalid tool schema for ${tool.name}`);
    }
  }
  console.log(`Self-test passed: ${tools.length} tools registered.`);
} else {
  startServer();
}
