export const DEFAULT_ACCOUNT_ID = "default";

export type MicrosoftGraphTokenPayload = {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  expires_on?: number | string;
  obtained_at?: number;
  scope?: string;
  token_type?: string;
  id_token_claims?: {
    preferred_username?: string;
    name?: string;
    oid?: string;
    tid?: string;
  };
};

export type OutlookAccountConfig = {
  enabled?: boolean;
  name?: string;
  tenantId?: string;
  clientId?: string;
  tokenFile?: string;
  defaultTo?: string;
  pollingIntervalMs?: number;
  webhookPublicBaseUrl?: string;
  webhookPath?: string;
};

export type ResolvedOutlookAccount = {
  accountId: string;
  enabled: boolean;
  name?: string;
  tenantId?: string;
  clientId?: string;
  tokenFile?: string;
  defaultTo?: string;
  pollingIntervalMs: number;
  webhookPublicBaseUrl?: string;
  webhookPath?: string;
  watchedFolderId?: string;
  subscriptionRenewBufferMs?: number;
};

export type GraphEmailAddress = {
  name?: string;
  address?: string;
};

export type GraphMessage = {
  id: string;
  conversationId?: string;
  subject?: string;
  bodyPreview?: string;
  receivedDateTime?: string;
  webLink?: string;
  from?: {
    emailAddress?: GraphEmailAddress;
  };
  sender?: {
    emailAddress?: GraphEmailAddress;
  };
  replyTo?: Array<{
    emailAddress?: GraphEmailAddress;
  }>;
  body?: {
    contentType?: string;
    content?: string;
  };
};

export type GraphSubscription = {
  id: string;
  resource: string;
  expirationDateTime: string;
  clientState?: string;
};
