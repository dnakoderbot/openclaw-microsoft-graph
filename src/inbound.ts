import {
  createReplyPrefixOptions,
  resolveInboundRouteEnvelopeBuilderWithRuntime,
} from "openclaw/plugin-sdk/compat";
import type { OpenClawConfig, PluginRuntime } from "openclaw/plugin-sdk/core";

import { fetchMessage, replyToMessage, sendMail } from "./graph-client.js";
import { resolveOutlookAccount } from "./config.js";
import type { GraphMessage, ResolvedOutlookAccount } from "./types.js";

const processedMessageIds = new Set<string>();

function messageSenderAddress(message: GraphMessage): string | undefined {
  return (
    message.replyTo?.[0]?.emailAddress?.address ??
    message.from?.emailAddress?.address ??
    message.sender?.emailAddress?.address
  );
}

function messageSenderName(message: GraphMessage): string | undefined {
  return (
    message.replyTo?.[0]?.emailAddress?.name ??
    message.from?.emailAddress?.name ??
    message.sender?.emailAddress?.name
  );
}

function stripHtml(input?: string): string {
  return (input ?? "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}

export async function processIncomingMessage(params: {
  runtime: PluginRuntime;
  cfg: OpenClawConfig;
  accountId: string;
  messageId: string;
}): Promise<void> {
  const { runtime, cfg, accountId, messageId } = params;
  if (processedMessageIds.has(messageId)) {
    return;
  }
  processedMessageIds.add(messageId);

  const account = resolveOutlookAccount(cfg, accountId);
  const message = await fetchMessage(account, messageId);
  const senderAddress = messageSenderAddress(message);
  if (!senderAddress) {
    runtime.system.enqueueSystemEvent(`Outlook inbound skipped: message ${messageId} has no sender.`);
    return;
  }

  const rawBody = stripHtml(message.body?.content) || message.bodyPreview || message.subject || "";
  const conversationId = message.conversationId || senderAddress.toLowerCase();
  const { route, buildEnvelope } = resolveInboundRouteEnvelopeBuilderWithRuntime({
    cfg,
    channel: "outlook",
    accountId: account.accountId,
    peer: {
      kind: "direct",
      id: conversationId,
    },
    runtime: runtime.channel,
    sessionStore: cfg.session?.store,
  });
  const fromLabel = messageSenderName(message) || senderAddress;
  const { storePath, body } = buildEnvelope({
    channel: "Outlook",
    from: fromLabel,
    timestamp: message.receivedDateTime ? Date.parse(message.receivedDateTime) : undefined,
    body: rawBody,
  });

  const ctxPayload = runtime.channel.reply.finalizeInboundContext({
    Body: body,
    BodyForAgent: rawBody,
    RawBody: rawBody,
    CommandBody: rawBody,
    From: `outlook:${senderAddress.toLowerCase()}`,
    To: `outlook:${account.defaultTo ?? "me"}`,
    SessionKey: route.sessionKey,
    AccountId: route.accountId,
    ChatType: "direct",
    ConversationLabel: fromLabel,
    SenderName: messageSenderName(message),
    SenderId: senderAddress.toLowerCase(),
    SenderUsername: senderAddress.toLowerCase(),
    Provider: "outlook",
    Surface: "outlook",
    MessageSid: message.id,
    MessageSidFull: message.id,
    ReplyToId: message.id,
    ReplyToIdFull: message.id,
    ThreadId: conversationId,
    OriginatingChannel: "outlook",
    OriginatingTo: `outlook:${account.defaultTo ?? "me"}`,
  });

  void runtime.channel.session
    .recordSessionMetaFromInbound({
      storePath,
      sessionKey: ctxPayload.SessionKey ?? route.sessionKey,
      ctx: ctxPayload,
    })
    .catch(() => {});

  const { onModelSelected, ...prefixOptions } = createReplyPrefixOptions({
    cfg,
    agentId: route.agentId,
    channel: "outlook",
    accountId: route.accountId,
  });

  await runtime.channel.reply.dispatchReplyWithBufferedBlockDispatcher({
    ctx: ctxPayload,
    cfg,
    dispatcherOptions: {
      ...prefixOptions,
      deliver: async (payload) => {
        const text = payload.text?.trim();
        if (!text) {
          return;
        }
        if (payload.replyToId) {
          await replyToMessage({
            account,
            messageId,
            text,
          });
          return;
        }
        await sendMail({
          account,
          to: senderAddress,
          subject: `Re: ${message.subject ?? "OpenClaw message"}`,
          text,
        });
      },
      onError: (error) => {
        runtime.system.enqueueSystemEvent(
          `Outlook reply failed for ${senderAddress}: ${String((error as Error)?.message ?? error)}`,
        );
      },
    },
    replyOptions: {
      onModelSelected,
    },
  });
}
