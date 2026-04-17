// Sentry MUST be initialized before any other module is required.
// This file is pre-loaded via `require("./instrument.js")` at the very top of server.js.
const Sentry = require("@sentry/node");
require("dotenv").config();

const dsn = process.env.SENTRY_DSN;

if (dsn) {
  Sentry.init({
    dsn,
    environment: process.env.SENTRY_ENVIRONMENT || process.env.NODE_ENV || "production",
    release: process.env.SENTRY_RELEASE || process.env.RAILWAY_GIT_COMMIT_SHA || undefined,
    tracesSampleRate: 0.2,
    sendDefaultPii: false,
    maxBreadcrumbs: 50,
    beforeSend(event) {
      if (event.request?.headers) {
        delete event.request.headers.authorization;
        delete event.request.headers.cookie;
        delete event.request.headers["x-api-key"];
      }
      if (event.extra) {
        for (const key of Object.keys(event.extra)) {
          if (/api[-_]?key|secret|token|password|ax_pass|client_secret|session_secret/i.test(key)) {
            event.extra[key] = "[REDACTED]";
          }
        }
      }
      return event;
    },
    beforeBreadcrumb(breadcrumb) {
      if (breadcrumb.category === "http" && breadcrumb.data?.url?.includes("/api/health")) return null;
      if (breadcrumb.category === "http" && breadcrumb.data?.url?.includes("/api/status")) return null;
      return breadcrumb;
    },
  });
  console.log("[Sentry] Initialized for environment:", process.env.NODE_ENV || "production");
} else {
  console.log("[Sentry] SENTRY_DSN not set — error tracking disabled");
}

module.exports = Sentry;
