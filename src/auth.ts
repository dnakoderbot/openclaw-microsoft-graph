import { readFile } from "node:fs/promises";

import type { ResolvedOutlookAccount, MicrosoftGraphTokenPayload } from "./types.js";

export async function loadTokenPayload(
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
