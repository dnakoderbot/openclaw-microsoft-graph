import type { IncomingMessage, ServerResponse } from "node:http";

import { getRuntimeApi } from "./runtime.js";
import { processIncomingMessage } from "./inbound.js";
import { resolveOutlookAccount } from "./config.js";

type GraphNotificationEnvelope = {
  value?: Array<{
    subscriptionId?: string;
    clientState?: string;
    resource?: string;
  }>;
};

function parseAccountId(clientState?: string): string {
  const prefix = "openclaw-outlook:";
  if (!clientState?.startsWith(prefix)) {
    return "default";
  }
  return clientState.slice(prefix.length) || "default";
}

function parseMessageId(resource?: string): string | null {
  if (!resource) {
    return null;
  }
  const match = resource.match(/messages\/([^/]+)$/i);
  return match?.[1] ?? null;
}

async function readJsonBody(req: IncomingMessage): Promise<GraphNotificationEnvelope> {
  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  const raw = Buffer.concat(chunks).toString("utf8").trim();
  if (!raw) {
    return {};
  }
  return JSON.parse(raw) as GraphNotificationEnvelope;
}

export async function handleOutlookWebhook(
  req: IncomingMessage,
  res: ServerResponse,
): Promise<boolean> {
  const url = new URL(req.url ?? "/", "http://localhost");
  if (url.searchParams.has("validationToken")) {
    res.statusCode = 200;
    res.setHeader("Content-Type", "text/plain");
    res.end(url.searchParams.get("validationToken") ?? "");
    return true;
  }

  if (req.method !== "POST") {
    return false;
  }

  const api = getRuntimeApi();
  const cfg = api.runtime.config.loadConfig();
  const envelope = await readJsonBody(req);
  res.statusCode = 202;
  res.end("accepted");

  for (const item of envelope.value ?? []) {
    const accountId = parseAccountId(item.clientState);
    const messageId = parseMessageId(item.resource);
    if (!messageId) {
      continue;
    }
    const account = resolveOutlookAccount(cfg, accountId);
    if (!account.enabled) {
      continue;
    }
    void processIncomingMessage({
      runtime: api.runtime,
      cfg,
      accountId,
      messageId,
    }).catch((error) => {
      api.logger.error(
        `Outlook inbound processing failed for ${accountId}/${messageId}: ${String(
          (error as Error)?.message ?? error,
        )}`,
      );
    });
  }

  return true;
}
