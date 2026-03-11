import type { ChannelPlugin, OpenClawConfig } from "openclaw/plugin-sdk/core";

import { fetchMe, sendMail } from "./graph-client.js";
import {
  listOutlookAccountIds,
  resolveDefaultOutlookAccountId,
  resolveOutlookAccount,
} from "./config.js";
import { ensureMailSubscription, stopMailSubscription } from "./subscriptions.js";
import type { ResolvedOutlookAccount } from "./types.js";

const meta = {
  id: "outlook",
  label: "Outlook",
  selectionLabel: "Outlook (Microsoft Graph)",
  docsPath: "/channels/outlook",
  docsLabel: "outlook",
  blurb: "Microsoft 365 email via delegated Microsoft Graph access.",
  aliases: ["microsoft-graph", "office365"],
  order: 58,
};

function waitForAbort(signal: AbortSignal): Promise<void> {
  if (signal.aborted) {
    return Promise.resolve();
  }
  return new Promise((resolve) => {
    signal.addEventListener("abort", () => resolve(), { once: true });
  });
}

function defaultSubject(text: string): string {
  const firstLine = text.split("\n").find((line) => line.trim());
  if (!firstLine) {
    return "OpenClaw message";
  }
  return firstLine.slice(0, 120);
}

export const outlookPlugin: ChannelPlugin<ResolvedOutlookAccount> = {
  id: "outlook",
  meta,
  capabilities: {
    chatTypes: ["direct", "thread"],
    threads: true,
    media: false,
    nativeCommands: false,
    blockStreaming: true,
  },
  reload: { configPrefixes: ["channels.outlook"] },
  config: {
    listAccountIds: (cfg) => listOutlookAccountIds(cfg),
    resolveAccount: (cfg, accountId) => resolveOutlookAccount(cfg, accountId),
    defaultAccountId: (_cfg) => resolveDefaultOutlookAccountId(),
    isConfigured: (account) => Boolean(account.tokenFile),
    describeAccount: (account) => ({
      accountId: account.accountId,
      name: account.name,
      enabled: account.enabled,
      configured: Boolean(account.tokenFile),
      tokenSource: account.tokenFile ? "file" : "none",
      webhookPath: account.webhookPath,
    }),
    resolveDefaultTo: ({ cfg, accountId }) => resolveOutlookAccount(cfg, accountId).defaultTo,
  },
  messaging: {
    normalizeTarget: (to) => to.trim().toLowerCase(),
    targetResolver: {
      looksLikeId: (value) => value.includes("@"),
      hint: "user@example.com",
    },
  },
  outbound: {
    deliveryMode: "direct",
    sendText: async ({ cfg, to, text, accountId }) => {
      const account = resolveOutlookAccount(cfg, accountId);
      await sendMail({
        account,
        to,
        subject: defaultSubject(text),
        text,
      });
      return {
        channel: "outlook",
        ok: true,
        id: `outlook:${Date.now()}`,
      };
    },
  },
  status: {
    defaultRuntime: {
      accountId: "default",
      running: false,
      lastStartAt: null,
      lastStopAt: null,
      lastError: null,
    },
    buildAccountSnapshot: async ({ account, runtime }) => {
      try {
        const me = await fetchMe(account);
        return {
          accountId: account.accountId,
          name: account.name ?? String(me.mail ?? me.userPrincipalName ?? ""),
          enabled: account.enabled,
          configured: Boolean(account.tokenFile),
          running: runtime?.running ?? false,
          connected: true,
          profile: {
            mail: me.mail,
            userPrincipalName: me.userPrincipalName,
            displayName: me.displayName,
          },
          tokenSource: account.tokenFile ? "file" : "none",
          webhookPath: account.webhookPath,
          webhookUrl: account.webhookPublicBaseUrl
            ? `${account.webhookPublicBaseUrl}${account.webhookPath}`
            : undefined,
          lastError: null,
        };
      } catch (error) {
        return {
          accountId: account.accountId,
          name: account.name,
          enabled: account.enabled,
          configured: Boolean(account.tokenFile),
          running: runtime?.running ?? false,
          connected: false,
          tokenSource: account.tokenFile ? "file" : "none",
          webhookPath: account.webhookPath,
          webhookUrl: account.webhookPublicBaseUrl
            ? `${account.webhookPublicBaseUrl}${account.webhookPath}`
            : undefined,
          lastError: String((error as Error)?.message ?? error),
        };
      }
    },
  },
  gateway: {
    startAccount: async (ctx) => {
      const me = await fetchMe(ctx.account);
      const subscription = await ensureMailSubscription({
        runtime: ctx.runtime,
        account: ctx.account,
      });
      ctx.log?.info?.(
        `[${ctx.accountId}] Outlook bootstrap connected as ${String(
          me.mail ?? me.userPrincipalName ?? "unknown",
        )}`,
      );
      if (subscription) {
        ctx.log?.info?.(
          `[${ctx.accountId}] Outlook subscription active until ${subscription.expirationDateTime}`,
        );
      }
      ctx.setStatus({
        ...ctx.getStatus(),
        running: true,
        connected: true,
        lastStartAt: Date.now(),
        lastError: null,
      });
      await waitForAbort(ctx.abortSignal);
      stopMailSubscription(ctx.accountId);
      ctx.setStatus({
        ...ctx.getStatus(),
        running: false,
        connected: false,
        lastStopAt: Date.now(),
      });
    },
  },
};
