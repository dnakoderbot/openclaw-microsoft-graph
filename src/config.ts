import type { OpenClawConfig } from "openclaw/plugin-sdk/core";

import { DEFAULT_ACCOUNT_ID, type OutlookAccountConfig, type ResolvedOutlookAccount } from "./types.js";

const DEFAULT_TOKEN_FILE = "/home/koderbot/.microsoft/tokens/msal_auth_result.json";
const DEFAULT_POLLING_INTERVAL_MS = 60_000;
const DEFAULT_WEBHOOK_PATH = "/plugins/outlook/webhook";
const DEFAULT_SUBSCRIPTION_RENEW_BUFFER_MS = 15 * 60_000;

function rawOutlookSection(cfg: OpenClawConfig): Record<string, unknown> {
  return ((cfg as Record<string, unknown>).channels as Record<string, unknown> | undefined)?.outlook as
    | Record<string, unknown>
    | undefined
    | Record<string, never> ?? {};
}

function readAccountConfig(cfg: OpenClawConfig, accountId?: string | null): OutlookAccountConfig {
  const section = rawOutlookSection(cfg);
  const normalizedAccountId = (accountId ?? DEFAULT_ACCOUNT_ID).trim() || DEFAULT_ACCOUNT_ID;
  if (normalizedAccountId === DEFAULT_ACCOUNT_ID) {
    return section as OutlookAccountConfig;
  }
  const accounts = section.accounts as Record<string, OutlookAccountConfig> | undefined;
  return accounts?.[normalizedAccountId] ?? {};
}

export function listOutlookAccountIds(cfg: OpenClawConfig): string[] {
  const section = rawOutlookSection(cfg);
  const accountIds = new Set<string>();
  accountIds.add(DEFAULT_ACCOUNT_ID);
  const accounts = section.accounts as Record<string, unknown> | undefined;
  for (const accountId of Object.keys(accounts ?? {})) {
    accountIds.add(accountId);
  }
  return [...accountIds];
}

export function resolveOutlookAccount(
  cfg: OpenClawConfig,
  accountId?: string | null,
): ResolvedOutlookAccount {
  const account = readAccountConfig(cfg, accountId);
  const normalizedAccountId = (accountId ?? DEFAULT_ACCOUNT_ID).trim() || DEFAULT_ACCOUNT_ID;
  return {
    accountId: normalizedAccountId,
    enabled: account.enabled !== false,
    name: account.name,
    tenantId: account.tenantId,
    clientId: account.clientId,
    tokenFile: account.tokenFile ?? DEFAULT_TOKEN_FILE,
    defaultTo: account.defaultTo,
    pollingIntervalMs: account.pollingIntervalMs ?? DEFAULT_POLLING_INTERVAL_MS,
    webhookPublicBaseUrl: account.webhookPublicBaseUrl,
    webhookPath: account.webhookPath ?? DEFAULT_WEBHOOK_PATH,
    watchedFolderId:
      typeof (account as Record<string, unknown>).watchedFolderId === "string"
        ? ((account as Record<string, unknown>).watchedFolderId as string)
        : "Inbox",
    subscriptionRenewBufferMs:
      typeof (account as Record<string, unknown>).subscriptionRenewBufferMs === "number"
        ? ((account as Record<string, unknown>).subscriptionRenewBufferMs as number)
        : DEFAULT_SUBSCRIPTION_RENEW_BUFFER_MS,
  };
}

export function resolveDefaultOutlookAccountId(): string {
  return DEFAULT_ACCOUNT_ID;
}
