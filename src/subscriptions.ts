import type { PluginRuntime } from "openclaw/plugin-sdk/core";

import { createMailSubscription, renewSubscription } from "./graph-client.js";
import type { GraphSubscription, ResolvedOutlookAccount } from "./types.js";

type ActiveSubscription = {
  accountId: string;
  subscription: GraphSubscription;
  renewTimer: NodeJS.Timeout | null;
};

const activeSubscriptions = new Map<string, ActiveSubscription>();

function computeRenewDelayMs(account: ResolvedOutlookAccount, expirationDateTime: string): number {
  const expiry = Date.parse(expirationDateTime);
  const delay = expiry - Date.now() - account.subscriptionRenewBufferMs;
  return Math.max(30_000, delay);
}

async function scheduleRenewal(
  runtime: PluginRuntime,
  account: ResolvedOutlookAccount,
  active: ActiveSubscription,
): Promise<void> {
  const delayMs = computeRenewDelayMs(account, active.subscription.expirationDateTime);
  if (active.renewTimer) {
    clearTimeout(active.renewTimer);
  }
  active.renewTimer = setTimeout(async () => {
    try {
      const renewed = await renewSubscription({
        account,
        subscriptionId: active.subscription.id,
        expiresAt: new Date(Date.now() + 60 * 60_000).toISOString(),
      });
      active.subscription = renewed;
      runtime.system.enqueueSystemEvent(
        `Outlook subscription renewed for ${account.accountId} until ${renewed.expirationDateTime}.`,
      );
      await scheduleRenewal(runtime, account, active);
    } catch (error) {
      runtime.system.enqueueSystemEvent(
        `Outlook subscription renewal failed for ${account.accountId}: ${String(
          (error as Error)?.message ?? error,
        )}`,
      );
    }
  }, delayMs);
}

export async function ensureMailSubscription(params: {
  runtime: PluginRuntime;
  account: ResolvedOutlookAccount;
}): Promise<GraphSubscription | null> {
  const { runtime, account } = params;
  if (!account.webhookPublicBaseUrl?.trim()) {
    return null;
  }
  const existing = activeSubscriptions.get(account.accountId);
  if (existing) {
    return existing.subscription;
  }
  const subscription = await createMailSubscription(account);
  const active: ActiveSubscription = {
    accountId: account.accountId,
    subscription,
    renewTimer: null,
  };
  activeSubscriptions.set(account.accountId, active);
  await scheduleRenewal(runtime, account, active);
  return subscription;
}

export function stopMailSubscription(accountId: string): void {
  const active = activeSubscriptions.get(accountId);
  if (!active) {
    return;
  }
  if (active.renewTimer) {
    clearTimeout(active.renewTimer);
  }
  activeSubscriptions.delete(accountId);
}
