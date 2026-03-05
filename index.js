import { logger } from "./logger.js";
import { streamManager } from "./stream-manager.js";
import { wecomChannelPlugin } from "./wecom/channel-plugin.js";
import { wecomHttpHandler } from "./wecom/http-handler.js";
import { responseUrls, setOpenclawConfig, setRuntime, streamMeta } from "./wecom/state.js";

// Periodic cleanup for streamMeta and expired responseUrls to prevent memory leaks.
setInterval(() => {
  const now = Date.now();
  // Clean streamMeta entries whose stream no longer exists in streamManager.
  for (const streamId of streamMeta.keys()) {
    if (!streamManager.hasStream(streamId)) {
      streamMeta.delete(streamId);
    }
  }
  // Clean expired responseUrls (older than 1 hour).
  for (const [key, entry] of responseUrls.entries()) {
    if (now > entry.expiresAt) {
      responseUrls.delete(key);
    }
  }
}, 60 * 1000).unref();

const plugin = {
  // Plugin id should match `openclaw.plugin.json` id (and config.plugins.entries key).
  id: "wecom",
  name: "Enterprise WeChat",
  description: "Enterprise WeChat AI Bot channel plugin for OpenClaw",
  configSchema: { type: "object", additionalProperties: false, properties: {} },
  register(api) {
    logger.info("WeCom plugin registering...");

    // Save runtime for message processing
    setRuntime(api.runtime);
    setOpenclawConfig(api.config);

    // Register channel
    api.registerChannel({ plugin: wecomChannelPlugin });
    logger.info("WeCom channel registered");

    // Register webhook HTTP handler so gateway routes /webhooks/* to this plugin.
    // WeCom callbacks use msg_signature verification, not Bearer-token auth.
    //
    // Strategy: register via both APIs for maximum compatibility.
    // - registerHttpRoute (OpenClaw 3.2+): prefix match catches /webhooks/*.
    // - registerHttpHandler (legacy): wildcard fallback if prefix match is
    //   unsupported (gateway does exact-path matching on older versions).
    if (typeof api.registerHttpRoute === "function") {
      api.registerHttpRoute({
        path: "/webhooks",
        handler: wecomHttpHandler,
        auth: "plugin",
        match: "prefix",
      });
      logger.info("WeCom HTTP route registered (auth: plugin, match: prefix)");
    }
    if (typeof api.registerHttpHandler === "function") {
      api.registerHttpHandler(wecomHttpHandler);
      logger.info("WeCom HTTP handler registered (legacy fallback)");
    }
  },
};

export default plugin;
export const register = (api) => plugin.register(api);
