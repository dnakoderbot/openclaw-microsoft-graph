import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

import {
  createReplyPrefixOptions,
  resolveInboundRouteEnvelopeBuilderWithRuntime,
} from "openclaw/plugin-sdk/compat";
import type { OpenClawConfig, PluginRuntime } from "openclaw/plugin-sdk/core";

import {
  fetchFileAttachments,
  fetchMessage,
  replyToMessage,
  sendMail,
  uploadFileToDrive,
} from "./graph-client.js";
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

function sanitizeFilename(input: string): string {
  return input.replace(/[^A-Za-z0-9._-]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 120) || "attachment";
}

async function materializeAttachments(params: {
  account: ResolvedOutlookAccount;
  message: GraphMessage;
}): Promise<{ local: string[]; drive: string[] }> {
  const { account, message } = params;
  if (!message.hasAttachments) {
    return { local: [], drive: [] };
  }
  const attachments = await fetchFileAttachments(account, message.id);
  if (!attachments.length) {
    return { local: [], drive: [] };
  }

  const targetDir = path.join(account.attachmentDownloadDir ?? "/tmp", sanitizeFilename(message.id));
  await mkdir(targetDir, { recursive: true });

  const saved: string[] = [];
  const mirrored: string[] = [];
  for (const attachment of attachments) {
    if (attachment.isInline) {
      continue;
    }
    const size = attachment.size ?? 0;
    if (size > (account.attachmentMaxBytes ?? 0)) {
      saved.push(
        `${attachment.name ?? "unnamed"} (skipped: ${size} bytes exceeds limit ${account.attachmentMaxBytes})`,
      );
      continue;
    }
    if (!attachment.contentBytes) {
      saved.push(`${attachment.name ?? "unnamed"} (metadata only; no contentBytes returned)`);
      continue;
    }
    const fileName = sanitizeFilename(attachment.name ?? `attachment-${saved.length + 1}`);
    const fullPath = path.join(targetDir, fileName);
    await writeFile(fullPath, Buffer.from(attachment.contentBytes, "base64"));
    saved.push(`${attachment.name ?? fileName} -> ${fullPath}`);
    if (account.mirrorInboundAttachmentsToDrive) {
      const remotePath = [
        account.driveWorkspacePath ?? "/OpenClaw",
        "inbound-attachments",
        sanitizeFilename(message.id),
        fileName,
      ].join("/");
      try {
        const driveUrl = await uploadFileToDrive({
          account,
          localPath: fullPath,
          remotePath,
        });
        mirrored.push(`${attachment.name ?? fileName} -> ${driveUrl}`);
      } catch (error) {
        mirrored.push(
          `${attachment.name ?? fileName} (Drive mirror skipped: ${String((error as Error)?.message ?? error)})`,
        );
      }
    }
  }

  return { local: saved, drive: mirrored };
}

function buildInboundBody(params: {
  message: GraphMessage;
  rawBody: string;
  savedAttachments: string[];
  mirroredAttachments: string[];
}): string {
  const lines = [params.rawBody.trim()];
  if (params.savedAttachments.length) {
    lines.push("");
    lines.push("Local attachments:");
    for (const item of params.savedAttachments) {
      lines.push(`- ${item}`);
    }
  }
  if (params.mirroredAttachments.length) {
    lines.push("");
    lines.push("OneDrive mirrors:");
    for (const item of params.mirroredAttachments) {
      lines.push(`- ${item}`);
    }
  }
  if (params.message.webLink) {
    lines.push("");
    lines.push(`Outlook link: ${params.message.webLink}`);
  }
  return lines.filter(Boolean).join("\n");
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
  const attachmentState = await materializeAttachments({ account, message });
  const bodyForAgent = buildInboundBody({
    message,
    rawBody,
    savedAttachments: attachmentState.local,
    mirroredAttachments: attachmentState.drive,
  });
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
    body: bodyForAgent,
  });

  const ctxPayload = runtime.channel.reply.finalizeInboundContext({
    Body: body,
    BodyForAgent: bodyForAgent,
    RawBody: bodyForAgent,
    CommandBody: bodyForAgent,
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
