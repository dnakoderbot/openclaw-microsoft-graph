import { chmod, readFile, writeFile } from "node:fs/promises";

import type { ResolvedOutlookAccount, MicrosoftGraphTokenPayload } from "./types.js";

const DEFAULT_SCOPE =
  "openid profile email offline_access User.Read Mail.ReadWrite Calendars.ReadWrite Contacts.ReadWrite Mail.Send Files.ReadWrite.All";
const REFRESH_BUFFER_MS = 5 * 60_000;

function resolveScope(payload: MicrosoftGraphTokenPayload): string {
  const raw = payload.scope?.trim() || DEFAULT_SCOPE;
  const scopes = new Set(raw.split(/\s+/).filter(Boolean));
  scopes.add("offline_access");
  return [...scopes].join(" ");
}

function parseExpiresAt(payload: MicrosoftGraphTokenPayload): number | null {
  if (typeof payload.expires_on === "number") {
    return payload.expires_on * 1000;
  }
  if (typeof payload.expires_on === "string" && payload.expires_on.trim()) {
    const asNumber = Number(payload.expires_on);
    if (Number.isFinite(asNumber)) {
      return asNumber * 1000;
    }
    const asDate = Date.parse(payload.expires_on);
    return Number.isFinite(asDate) ? asDate : null;
  }
  if (typeof payload.obtained_at === "number" && typeof payload.expires_in === "number") {
    return payload.obtained_at + payload.expires_in * 1000;
  }
  return null;
}

function shouldRefreshToken(payload: MicrosoftGraphTokenPayload): boolean {
  const expiresAt = parseExpiresAt(payload);
  if (expiresAt === null) {
    return false;
  }
  return Date.now() >= expiresAt - REFRESH_BUFFER_MS;
}

async function readTokenPayload(
  account: ResolvedOutlookAccount,
): Promise<MicrosoftGraphTokenPayload> {
  const tokenFile = account.tokenFile?.trim();
  if (!tokenFile) {
    throw new Error("Outlook tokenFile is not configured");
  }
  const raw = await readFile(tokenFile, "utf8");
  const payload = JSON.parse(raw) as MicrosoftGraphTokenPayload;
  if (!payload.access_token?.trim()) {
    throw new Error(`Outlook token file does not contain access_token: ${tokenFile}`);
  }
  return payload;
}

async function persistTokenPayload(
  account: ResolvedOutlookAccount,
  payload: MicrosoftGraphTokenPayload,
): Promise<void> {
  const tokenFile = account.tokenFile?.trim();
  if (!tokenFile) {
    throw new Error("Outlook tokenFile is not configured");
  }
  await writeFile(tokenFile, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  await chmod(tokenFile, 0o600);
}

export async function refreshTokenPayload(
  account: ResolvedOutlookAccount,
  current?: MicrosoftGraphTokenPayload,
): Promise<MicrosoftGraphTokenPayload> {
  const payload = current ?? (await readTokenPayload(account));
  const refreshToken = payload.refresh_token?.trim();
  if (!refreshToken) {
    throw new Error("Outlook token file does not contain refresh_token");
  }
  if (!account.tenantId?.trim()) {
    throw new Error("Outlook tenantId is required for token refresh");
  }
  if (!account.clientId?.trim()) {
    throw new Error("Outlook clientId is required for token refresh");
  }

  const body = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: account.clientId,
    refresh_token: refreshToken,
    scope: resolveScope(payload),
  });
  const response = await fetch(
    `https://login.microsoftonline.com/${account.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body,
    },
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Microsoft token refresh failed: HTTP ${response.status} ${errorText}`);
  }

  const refreshed = (await response.json()) as MicrosoftGraphTokenPayload;
  if (!refreshed.access_token?.trim()) {
    throw new Error("Microsoft token refresh response did not contain access_token");
  }

  const merged: MicrosoftGraphTokenPayload = {
    ...payload,
    ...refreshed,
    refresh_token: refreshed.refresh_token ?? payload.refresh_token,
    obtained_at: Date.now(),
  };

  await persistTokenPayload(account, merged);
  return merged;
}

export async function loadTokenPayload(
  account: ResolvedOutlookAccount,
): Promise<MicrosoftGraphTokenPayload> {
  const payload = await readTokenPayload(account);
  if (shouldRefreshToken(payload)) {
    return refreshTokenPayload(account, payload);
  }
  return payload;
}
