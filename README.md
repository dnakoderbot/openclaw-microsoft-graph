# OpenClaw Microsoft Graph

Standalone OpenClaw plugin project for Microsoft Graph integration.

License: MIT

## What This Plugin Does

This plugin is intended to let OpenClaw operate as a real Microsoft 365 user through delegated Microsoft Graph access.

Current live scope:

- Outlook email as a real inbound/outbound OpenClaw channel
- delegated token loading from a local token file
- refresh-token based Graph access refresh
- Outlook inbox webhook subscription bootstrap and renewal
- inbound email fetch and session dispatch into OpenClaw
- outbound email reply/send through Graph
- inbound attachment staging to local disk
- optional OneDrive mirroring for inbound attachments

Implemented Graph groundwork:

- mail
- calendar read/create helpers
- drive upload helper
- contacts/mail/calendar/file delegated auth model

## Current Runtime Model

This plugin is currently optimized for a self-hosted OpenClaw deployment where:

- OpenClaw runs on a VPS or workstation
- Microsoft Graph delegated auth is already established outside the plugin
- the plugin reads a local token file produced by a Microsoft OAuth flow
- a public HTTPS endpoint exists for Graph change notifications

Today, the plugin uses a local token file. The next auth milestone is first-class provider login/import so OpenClaw can own the Microsoft login lifecycle directly.

## Outlook Channel Behavior

Current webhook behavior:

- `GET /plugins/outlook/webhook?validationToken=...` echoes the token for Graph subscription validation
- `POST /plugins/outlook/webhook` accepts Graph notifications, fetches the message from Graph, and dispatches it into OpenClaw
- replies use Graph mail reply/send APIs

Inbound message behavior:

- each incoming message is fetched from Graph
- message text is normalized into an agent-visible body
- file attachments are downloaded locally
- downloaded attachments can be mirrored into OneDrive
- the agent sees attachment paths and, when available, mirrored OneDrive URLs in the inbound message context

## Config

OpenClaw channel config keys:

```json
{
  "channels": {
    "outlook": {
      "enabled": true,
      "name": "DNA Koderbot Outlook",
      "tenantId": "YOUR_TENANT_ID",
      "clientId": "YOUR_CLIENT_ID",
      "defaultTo": "koderbot@dnakode.com",
      "tokenFile": "/home/koderbot/.microsoft/tokens/msal_auth_result.json",
      "webhookPublicBaseUrl": "https://your-public-host.example.com",
      "webhookPath": "/plugins/outlook/webhook",
      "watchedFolderId": "Inbox",
      "attachmentDownloadDir": "/home/koderbot/.openclaw/outlook-attachments",
      "attachmentMaxBytes": 10485760,
      "driveWorkspacePath": "/OpenClaw",
      "mirrorInboundAttachmentsToDrive": true,
      "driveSimpleUploadMaxBytes": 4194304
    }
  }
}
```

Field notes:

- `tenantId` and `clientId` are required for Microsoft refresh-token exchange
- `tokenFile` points to a JSON token payload with `access_token` and `refresh_token`
- `webhookPublicBaseUrl` must be public HTTPS for Graph subscriptions
- `attachmentDownloadDir` is the local staging area for inbound attachments
- `driveWorkspacePath` is the root OneDrive path used for bot artifacts
- `driveSimpleUploadMaxBytes` uses Graph simple upload and is intentionally conservative

## Microsoft 365 Permissions

Current expected delegated scopes:

- `openid`
- `profile`
- `email`
- `offline_access`
- `User.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `Calendars.ReadWrite`
- `Contacts.ReadWrite`
- `Files.ReadWrite.All`

## What OpenClaw Can Use From Microsoft

What is already useful today:

- Outlook inbox as a primary inbound channel
- Outlook replies and direct sends
- attachment-aware inbound tasks
- OneDrive as a place to mirror inbound artifacts

What this Graph foundation can support next:

- calendar reading, scheduling, and invite handling
- contacts resolution and contact updates
- richer OneDrive and SharePoint workspace behavior
- message search, drafting, filing, and mailbox actions

## Repository Layout

```text
src/
  auth.ts
  calendar.ts
  channel.ts
  config.ts
  graph-client.ts
  inbound.ts
  index.ts
  runtime.ts
  subscriptions.ts
  types.ts
  webhook.ts
```

Module overview:

- `auth.ts`: token loading and refresh
- `channel.ts`: OpenClaw channel definition and lifecycle
- `graph-client.ts`: Microsoft Graph mail/drive/subscription primitives
- `calendar.ts`: Microsoft Graph calendar helpers
- `inbound.ts`: inbound email processing and attachment staging
- `subscriptions.ts`: Outlook webhook subscription lifecycle
- `webhook.ts`: Graph webhook route handler

## Operational Notes

- The plugin assumes delegated user access, not app-only service access.
- Microsoft webhook subscriptions expire and are renewed at runtime.
- Outlook is working independently from the Anthropic model-auth issue; mail transport and Graph access are separate from model inference credentials.
- On a reverse-proxied deployment, OpenClaw should trust the local proxy and explicitly allow the public Control UI origin.

## Roadmap

Near-term:

1. First-class Microsoft auth/provider import instead of token-file bootstrap
2. Thread/session mapping hardening for Outlook conversations
3. Calendar integration into OpenClaw workflows
4. OneDrive workspace publishing for generated artifacts
5. Contacts and directory enrichment

Longer-term:

1. Shared mailbox and shared calendar support
2. SharePoint library publishing
3. Rich attachment outbound flows
4. Draft/review/send mail workflows
