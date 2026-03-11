import type {
  GraphMessage,
  GraphSubscription,
  ResolvedOutlookAccount,
} from "./types.js";
import { loadTokenPayload } from "./auth.js";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

export async function graphFetch(
  account: ResolvedOutlookAccount,
  path: string,
  init?: RequestInit,
): Promise<Response> {
  const token = await loadTokenPayload(account);
  const headers = new Headers(init?.headers ?? {});
  headers.set("Authorization", `Bearer ${token.access_token}`);
  if (!headers.has("Content-Type") && init?.body) {
    headers.set("Content-Type", "application/json");
  }
  return fetch(`${GRAPH_ROOT}${path}`, { ...init, headers });
}

export async function fetchMe(account: ResolvedOutlookAccount): Promise<Record<string, unknown>> {
  const response = await graphFetch(account, "/me");
  if (!response.ok) {
    throw new Error(`Graph /me failed: HTTP ${response.status}`);
  }
  return (await response.json()) as Record<string, unknown>;
}

export async function sendMail(params: {
  account: ResolvedOutlookAccount;
  to: string;
  subject: string;
  text: string;
}): Promise<void> {
  const response = await graphFetch(params.account, "/me/sendMail", {
    method: "POST",
    body: JSON.stringify({
      message: {
        subject: params.subject,
        body: {
          contentType: "Text",
          content: params.text,
        },
        toRecipients: [
          {
            emailAddress: {
              address: params.to,
            },
          },
        ],
      },
      saveToSentItems: true,
    }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph sendMail failed: HTTP ${response.status} ${body}`);
  }
}

export async function replyToMessage(params: {
  account: ResolvedOutlookAccount;
  messageId: string;
  text: string;
}): Promise<void> {
  const response = await graphFetch(params.account, `/me/messages/${params.messageId}/reply`, {
    method: "POST",
    body: JSON.stringify({
      comment: params.text,
    }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph reply failed: HTTP ${response.status} ${body}`);
  }
}

export async function fetchMessage(
  account: ResolvedOutlookAccount,
  messageId: string,
): Promise<GraphMessage> {
  const response = await graphFetch(
    account,
    `/me/messages/${encodeURIComponent(messageId)}?$select=id,conversationId,subject,bodyPreview,receivedDateTime,webLink,from,sender,replyTo,body`,
  );
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph fetch message failed: HTTP ${response.status} ${body}`);
  }
  return (await response.json()) as GraphMessage;
}

export async function createMailSubscription(
  account: ResolvedOutlookAccount,
): Promise<GraphSubscription> {
  const resource =
    account.watchedFolderId && account.watchedFolderId !== "Inbox"
      ? `/me/mailFolders('${account.watchedFolderId}')/messages`
      : "/me/mailFolders('Inbox')/messages";
  const notificationUrl = `${String(account.webhookPublicBaseUrl ?? "").replace(/\/$/, "")}${account.webhookPath}`;
  const expiration = new Date(Date.now() + 60 * 60_000).toISOString();
  const response = await graphFetch(account, "/subscriptions", {
    method: "POST",
    body: JSON.stringify({
      changeType: "created",
      notificationUrl,
      resource,
      expirationDateTime: expiration,
      clientState: `openclaw-outlook:${account.accountId}`,
    }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph create subscription failed: HTTP ${response.status} ${body}`);
  }
  return (await response.json()) as GraphSubscription;
}

export async function renewSubscription(params: {
  account: ResolvedOutlookAccount;
  subscriptionId: string;
  expiresAt: string;
}): Promise<GraphSubscription> {
  const response = await graphFetch(params.account, `/subscriptions/${params.subscriptionId}`, {
    method: "PATCH",
    body: JSON.stringify({
      expirationDateTime: params.expiresAt,
    }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph renew subscription failed: HTTP ${response.status} ${body}`);
  }
  return (await response.json()) as GraphSubscription;
}
