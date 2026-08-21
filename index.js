/**
 * utm-gsheets — Cordis plugin for DeepSeek Harness.
 *
 * Registers the gsheets_* tool family (meta, list, write, append, create,
 * clear) for working with Google Sheets through the official v4 API. Reuses
 * the user's existing OAuth token (path set at deploy time via config.token_path
 * or the GSHEETS_TOKEN_PATH env var) with automatic refresh via the refresh
 * token — no service account, no Google Cloud setup needed.
 *
 * Install: `pnpm add "file:<plugins-dir>/gsheets"` in the dsh profile, and add
 * to `cordis.patch.yml` — MUST be an insert list:
 *
 *   - insert:
 *       - id: utm-gsheets
 *         name: utm-gsheets
 *         config:
 *           token_path: "<путь-к-google_token.json>"
 *
 * The token file itself is never in the repo — only the path is configured,
 * and no machine-specific paths/logins are hardcoded.
 */

import { defineTool } from '@deepseek-ai/dsh-tools'
import { buildConfig } from './src/config.js'
import { makeAuthenticatedClient } from './src/client.js'
import { buildTools } from './src/tools.js'

export const name = 'utm-gsheets'
export const inject = ['tools']

/** Permissive output schema: our tools return heterogeneous result objects. */
const OUTPUT_SCHEMA = { type: 'object', additionalProperties: true }

const MAX_RENDER_CHARS = 30_000

function renderResult(_args, value) {
  if (value && typeof value.ok === 'boolean' && value.ok === false) {
    return [{ type: 'text', text: `gsheets error: ${value.error}` }]
  }
  let json
  try {
    json = JSON.stringify(value, null, 2)
  } catch {
    json = String(value)
  }
  if (json.length > MAX_RENDER_CHARS) {
    json = json.slice(0, MAX_RENDER_CHARS) + '\n… (truncated)'
  }
  return [{ type: 'text', text: json }]
}

/** Build a fresh authenticated client from the plugin config. */
function hydrate(config) {
  const cfg = buildConfig(config)
  return makeAuthenticatedClient(null, cfg)
}

/** Plugin entry point: register one tool per definition. */
export function apply(ctx, config = {}) {
  for (const def of buildTools(() => hydrate(config))) {
    ctx.tools.register(defineTool({
      name: def.name,
      description: def.description,
      parameters: def.parameters,
      output: { schema: OUTPUT_SCHEMA, render: renderResult },
      execute: async (args) => {
        try {
          return await def.handler(args ?? {})
        } catch (e) {
          return { ok: false, error: e?.message ?? String(e) }
        }
      },
    }))
  }
}