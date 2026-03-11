# OpenClaw Microsoft Graph

Standalone OpenClaw plugin project for Microsoft Graph integration.

License: MIT

Current scope:

- Outlook email as a channel
- delegated token loading from a local token file
- Microsoft Graph client wrapper
- channel scaffold for outbound and inbound subscription support

OpenClaw channel config keys:

```json
{
  "channels": {
    "outlook": {
      "enabled": true,
      "name": "DNA Koderbot Outlook",
      "defaultTo": "koderbot@dnakode.com",
      "tokenFile": "/home/koderbot/.microsoft/tokens/msal_auth_result.json",
      "webhookPublicBaseUrl": "https://your-public-host.example.com",
      "webhookPath": "/plugins/outlook/webhook",
      "watchedFolderId": "Inbox",
      "attachmentDownloadDir": "/home/koderbot/.openclaw/outlook-attachments",
      "attachmentMaxBytes": 10485760
    }
  }
}
```

Current webhook behavior:

- `GET /plugins/outlook/webhook?validationToken=...` echoes the token for Graph subscription validation
- `POST /plugins/outlook/webhook` accepts Graph notifications, fetches the message from Graph, and dispatches it into OpenClaw
- inbound file attachments are downloaded locally and appended to the agent-visible message body
- replies use Graph mail reply/send APIs

Planned next steps:

1. Replace token-file bootstrap with first-class plugin login flow
2. Add outbound email send/reply
3. Add inbox subscriptions + webhook notifications
4. Map email threads to OpenClaw sessions
5. Extend to calendar, contacts, and files
