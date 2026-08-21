/**
 * Plugin config: where the Google OAuth token lives.
 *
 * The token file itself is NOT in the repo — it lives in the user's secret
 * store. The plugin only stores the *path* (or client id / secret inline) in
 * its cordis.patch.yml config, so no credential material is checked into
 * source and no machine-specific paths are hardcoded. The token path is
 * provided at deploy time via config.token_path or the GSHEETS_TOKEN_PATH env
 * variable.
 */

/** Merge raw plugin config with process.env fallbacks. */
export function buildConfig(config = {}) {
  const c = config ?? {}
  return {
    token_path:
      c.token_path ||
      process.env.GSHEETS_TOKEN_PATH ||
      null,
    client_id: c.client_id || process.env.GSHEETS_CLIENT_ID || null,
    client_secret: c.client_secret || process.env.GSHEETS_CLIENT_SECRET || null,
    // scopes are already baked into the refresh token; kept for clarity
    scopes: c.scopes ?? null,
  }
}