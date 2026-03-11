import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";

import { outlookPlugin } from "./channel.js";
import { setRuntimeApi } from "./runtime.js";
import { handleOutlookWebhook } from "./webhook.js";

const plugin = {
  id: "microsoft-graph",
  name: "Microsoft Graph",
  description: "OpenClaw Outlook channel plugin backed by Microsoft Graph",
  configSchema: {
    type: "object",
    additionalProperties: false,
    properties: {},
  },
  register(api: OpenClawPluginApi) {
    setRuntimeApi(api);
    api.registerChannel({ plugin: outlookPlugin });
    api.registerHttpRoute({
      path: "/plugins/outlook/webhook",
      auth: "plugin",
      match: "exact",
      handler: handleOutlookWebhook,
    });
  },
};

export default plugin;
