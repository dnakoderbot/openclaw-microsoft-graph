import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";

let runtimeApi: OpenClawPluginApi | null = null;

export function setRuntimeApi(api: OpenClawPluginApi): void {
  runtimeApi = api;
}

export function getRuntimeApi(): OpenClawPluginApi {
  if (!runtimeApi) {
    throw new Error("Microsoft Graph plugin runtime is not initialized");
  }
  return runtimeApi;
}
