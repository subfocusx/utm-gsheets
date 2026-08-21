/**
 * Google OAuth2 client + Sheets API wrapper.
 *
 * Uses the installed-app OAuth token that already exists on the host machine.
 * The plugin holds the token path / client id / secret in its config (or env
 * vars) so credentials never live in the repo — and no machine-specific paths
 * are hardcoded. The access token is refreshed automatically using the refresh
 * token whenever it has expired or the API replies 401.
 *
 * Environment variables are a fallback for all fields, matching the pattern
 * used by the yandex-api plugin (GSHEETS_*).
 */

import { readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'

/** No default path on purpose: the token location is deploy-time config. */
export const DEFAULT_TOKEN_PATH = null

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets'
const OAUTH_TOKEN_URI = 'https://oauth2.googleapis.com/token'

// ---------------------------------------------------------------------------
// ID / range helpers
// ---------------------------------------------------------------------------

/** Normalize a spreadsheet reference: URL or bare spreadsheet id. */
export function extractSpreadsheetId(input) {
  const s = String(input ?? '').trim()
  if (!s) throw new Error('не передан spreadsheet_id / URL таблицы')
  const m = /\/spreadsheets\/d\/([^/?#]+)/.exec(s)
  if (m) return m[1]
  if (s.includes('/')) throw new Error(`не удалось распознать spreadsheet id из: ${s}`)
  return s
}

/** 0-based column index -> A1 letter(s): 0->A, 25->Z, 26->AA. */
export function columnLetter(idx) {
  let out = ''
  let n = idx
  while (n >= 0) {
    out = String.fromCharCode(65 + (n % 26)) + out
    n = Math.floor(n / 26) - 1
  }
  return out
}

/**
 * A1 range for a rectangle of values.
 * @param sheetName sheet/tab title or null for the default sheet
 * @param startRow 0-based
 * @param startCol 0-based
 * @param nRows rows count (>=1)
 * @param nCols cols count (>=1)
 */
export function a1Range(sheetName, startRow, startCol, nRows, nCols) {
  const esc = (t) => (/\s/.test(t) ? `'${t.replace(/'/g, "''")}'` : t)
  const col1 = columnLetter(startCol)
  const col2 = columnLetter(startCol + Math.max(nCols, 1) - 1)
  const row1 = startRow + 1
  const row2 = startRow + Math.max(nRows, 1)
  const sheet = sheetName != null && sheetName !== '' ? esc(sheetName) + '!' : ''
  return `${sheet}${col1}${row1}:${col2}${row2}`
}

/** "A1" -> [col0, row0] (0-based). */
export function parseCell(cell) {
  const m = /^([A-Za-z]+)(\d+)$/.exec(String(cell ?? '').trim())
  if (!m) throw new Error(`неверная ссылка на ячейку: ${cell} (ожидается вида A1)`)
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return [col - 1, Number(m[2]) - 1]
}

/**
 * "A1" or "A1:C3" -> {startRow,endRow,startCol,endCol} (0-based inclusive).
 * A bare "A1" expands to a single cell.
 */
export function parseA1Range(range) {
  const s = String(range ?? '').trim()
  if (!s) throw new Error('не передан range (A1-вид)')
  const [a, b] = s.includes(':') ? s.split(':') : [s, s]
  const [sc, sr] = parseCell(a)
  const [ec, er] = parseCell(b)
  return { startRow: sr, endRow: er, startCol: sc, endCol: ec }
}

/** "100px" / "100" -> 100. Throws unless positive. */
export function parseSize(value, label = 'размер') {
  const n = Number(String(value ?? '').replace(/px$/i, '').trim())
  if (!Number.isFinite(n) || n <= 0) throw new Error(`${label}: ожидается число > 0 (например 100 или 100px)`)
  return Math.round(n)
}

/**
 * Parse a key like "C" (column letter), "A:Z" (column span), "5" (row),
 * "1:5" (row span), or a plain A1 range. Returns a normalized grid object
 * { sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex }.
 * Accepts an axis hint when the key is ambiguous (numeric key is a row).
 */
export function rangeObject(sheetId, key) {
  const k = String(key ?? '').trim().toUpperCase()
  if (!k) throw new Error('не передан диапазон')
  const colRx = /^[A-Z]+(?::[A-Z]+)?$/
  const rowRx = /^\d+(?::\d+)?$/
  if (colRx.test(k)) {
    const [c1, c2] = k.split(':').map((c) => parseCell(c + '1')[0])
    return { sheetId, startRowIndex: 0, endRowIndex: 1048576, startColumnIndex: c1, endColumnIndex: (c2 ?? c1) + 1 }
  }
  if (rowRx.test(k)) {
    const [r1, r2] = k.split(':').map((r) => Number(r))
    return { sheetId, startRowIndex: r1 - 1, endRowIndex: r2 ?? r1, startColumnIndex: 0, endColumnIndex: 18278 }
  }
  // plain A1 range
  const { startRow, endRow, startCol, endCol } = parseA1Range(k)
  return { sheetId, startRowIndex: startRow, endRowIndex: endRow + 1, startColumnIndex: startCol, endColumnIndex: endCol + 1 }
}

/** Hex "#RRGGBB" or "RRGGBB" -> {red,green,blue} for Sheets API. */
export function hexToRgb(hex) {
  const h = String(hex ?? '').replace(/^#/, '').trim()
  if (!/^[0-9a-fA-F]{6}$/.test(h)) throw new Error(`неверный цвет: ${hex} (ожидается #RRGGBB)`)
  return {
    red: parseInt(h.slice(0, 2), 16),
    green: parseInt(h.slice(2, 4), 16),
    blue: parseInt(h.slice(4, 6), 16),
  }
}

/** Map a JS value to a Sheets cell userEnteredValue object. */
export function toCellValue(v) {
  if (typeof v === 'number') return { numberValue: v }
  if (typeof v === 'boolean') return { boolValue: v }
  return { stringValue: String(v ?? '') }
}

// ---------------------------------------------------------------------------
// BatchUpdate request builders
// ---------------------------------------------------------------------------

/** repeatCell: apply formatting over an A1 range.
 *  format keys: bold, italic, fontSize, color(hex), fill(hex),
 *  horizontal, vertical (alignment). */
export function repeatCellRequest(sheetId, a1, format = {}) {
  const uf = {}
  const fields = []
  if (format.bold != null) {
    uf.textFormat = { ...(uf.textFormat ?? {}), bold: Boolean(format.bold) }
    fields.push('userEnteredFormat.textFormat.bold')
  }
  if (format.italic != null) {
    uf.textFormat = { ...(uf.textFormat ?? {}), italic: Boolean(format.italic) }
    fields.push('userEnteredFormat.textFormat.italic')
  }
  if (format.fontSize != null) {
    uf.textFormat = { ...(uf.textFormat ?? {}), fontSize: parseSize(format.fontSize) }
    fields.push('userEnteredFormat.textFormat.fontSize')
  }
  if (format.color) {
    uf.textFormat = { ...(uf.textFormat ?? {}), foregroundColor: hexToRgb(format.color) }
    fields.push('userEnteredFormat.textFormat.foregroundColor')
  }
  if (format.fill) {
    uf.backgroundColor = hexToRgb(format.fill)
    fields.push('userEnteredFormat.backgroundColor')
  }
  if (format.horizontal) {
    uf.horizontalAlignment = format.horizontal
    fields.push('userEnteredFormat.horizontalAlignment')
  }
  if (format.vertical) {
    uf.verticalAlignment = format.vertical
    fields.push('userEnteredFormat.verticalAlignment')
  }
  return {
    repeatCell: {
      range: rangeObject(sheetId, a1),
      cell: { userEnteredFormat: uf },
      fields: fields.join(','),
    },
  }
}

/** setBasicFilter over a used A1 range. */
export function basicFilterRequest(sheetId, a1) {
  return { setBasicFilter: { filter: { range: rangeObject(sheetId, a1) } } }
}

/** updateSheetProperties: freeze rows/cols. */
export function freezeRequest(sheetId, rows = 0, cols = 0) {
  return {
    updateSheetProperties: {
      properties: {
        sheetId,
        gridProperties: {
          frozenRowCount: Math.max(0, Number(rows) || 0),
          frozenColumnCount: Math.max(0, Number(cols) || 0),
        },
      },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    },
  }
}

/** updateDimensionProperties: set row heights / column widths.
 *  dimension: 'ROWS' | 'COLUMNS', key like "A:Z" or "1:5" or plain. */
export function dimensionSizeRequest(sheetId, dimension, key, valueExpr) {
  const r = rangeObject(sheetId, key)
  return {
    updateDimensionProperties: {
      range: {
        sheetId,
        dimension,
        startIndex: dimension === 'COLUMNS' ? r.startColumnIndex : r.startRowIndex,
        endIndex: dimension === 'COLUMNS' ? r.endColumnIndex : r.endRowIndex,
      },
      properties: { pixelSize: parseSize(valueExpr) },
      fields: 'pixelSize',
    },
  }
}

/** updateCells: write rows/values grid to an A1 range (ROWS major). */
export function updateCellsRequest(sheetId, a1, values) {
  const grid = (Array.isArray(values) ? values : [values])
    .map((row) => (Array.isArray(row) ? row : [row]).map((v) => ({ userEnteredValue: toCellValue(v) })))
  return {
    updateCells: {
      range: rangeObject(sheetId, a1),
      rows: grid,
      fields: 'userEnteredValue',
    },
  }
}

// ---------------------------------------------------------------------------
// Token client
// ---------------------------------------------------------------------------

function resolveTokenPath(configured, config) {
  const fromCfg = configured
  const fromConf = config?.token_path
  const fromEnv = process.env.GSHEETS_TOKEN_PATH
  return fromCfg || fromConf || fromEnv || DEFAULT_TOKEN_PATH
}

function requireTokenPath(path) {
  if (!path) {
    throw new Error(
      'путь к токену не задан. Укажите config.token_path в cordis.patch.yml ' +
      'или переменную GSHEETS_TOKEN_PATH',
    )
  }
  return path
}

/**
 * Load + validate the OAuth token file. Throws with a clear message if the
 * refresh path is impossible. The token file is read once at client build;
 * the access token is refreshed lazily (see sheetsRequest).
 */
export async function makeAuthenticatedClient(tokenPath = null, config = {}) {
  const path = requireTokenPath(resolveTokenPath(tokenPath, config))
  if (!existsSync(path)) {
    throw new Error(
      `токен не найден: ${path}. Укажите путь в конфиге плагина (config.token_path) ` +
      `или в переменной GSHEETS_TOKEN_PATH`,
    )
  }
  const raw = JSON.parse(await readFile(path, 'utf8'))
  const missing = []
  if (!raw.refresh_token) missing.push('refresh_token')
  if (!raw.client_id) missing.push('client_id')
  if (!raw.client_secret) missing.push('client_secret')
  if (missing.length) {
    throw new Error(
      `в токене ${path} отсутствуют: ${missing.join(', ')} — нужен refresh_token + ` +
      `client_id + client_secret (installed-app OAuth) для обновления access-токена`,
    )
  }
  return {
    access_token: raw.token || raw.access_token || null,
    refresh_token: raw.refresh_token,
    client_id: raw.client_id,
    client_secret: raw.client_secret,
    token_uri: raw.token_uri || OAUTH_TOKEN_URI,
    _expiry: raw.expiry ? new Date(raw.expiry).getTime() : null,
  }
}

/** True when the cached access token is expired (30s safety margin). */
function isExpired(client) {
  if (!client._expiry) return true // unknown -> refresh
  return client._expiry - 30_000 < Date.now()
}

/** Reuse or refresh (if expired/missing) the access token. */
async function ensureAccessToken(client) {
  if (client.access_token && !isExpired(client)) {
    return client.access_token
  }
  const params = new URLSearchParams({
    client_id: client.client_id,
    client_secret: client.client_secret,
    refresh_token: client.refresh_token,
    grant_type: 'refresh_token',
  })
  const res = await fetch(client.token_uri, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  })
  if (!res.ok) {
    const text = await res.text()
    throw new Error(`Google OAuth refresh failed (${res.status}): ${text.slice(0, 300)}`)
  }
  const data = await res.json()
  client.access_token = data.access_token
  client._expiry = Date.now() + (data.expires_in || 3600) * 1000
  return client.access_token
}

async function rawRequest(client, method, url, body) {
  const token = await ensureAccessToken(client)
  const headers = { Authorization: `Bearer ${token}` }
  if (body !== null && body !== undefined) headers['Content-Type'] = 'application/json'
  return fetch(url, {
    method,
    headers,
    body: body !== null && body !== undefined ? JSON.stringify(body) : undefined,
  })
}

/**
 * Authorized request to the Google Sheets v4 API. On a 401, forces a token
 * refresh and retries once (handles a token that expired mid-flight).
 * Returns parsed JSON on success; throws with the API error message otherwise.
 */
export async function sheetsRequest(client, method, url, body = null) {
  let res = await rawRequest(client, method, url, body)
  if (res.status === 401 && client.refresh_token) {
    client.access_token = null
    client._expiry = null
    res = await rawRequest(client, method, url, body)
  }
  const text = await res.text()
  let data = null
  if (text) {
    try {
      data = JSON.parse(text)
    } catch {
      data = { raw: text.slice(0, 500) }
    }
  }
  if (!res.ok) {
    const msg = data?.error?.message || data?.raw || text.slice(0, 300)
    const code = data?.error?.code ?? res.status
    throw new Error(`Google Sheets API ${code}: ${msg}`)
  }
  return data
}

/** POST /spreadsheets/{id}:batchUpdate with a requests array. */
export function batchUpdate(client, spreadsheetId, requests) {
  return sheetsRequest(client, 'POST', `${SHEETS_API}/${spreadsheetId}:batchUpdate`, { requests })
}

/** GET /spreadsheets/{id}/values:batchGet for several A1 ranges (ROWS planned). */
export function batchGetValues(client, spreadsheetId, ranges) {
  const qs = new URLSearchParams()
  for (const r of ranges) qs.append('ranges', r)
  qs.set('majorDimension', 'ROWS')
  return sheetsRequest(client, 'GET', `${SHEETS_API}/${spreadsheetId}/values:batchGet?${qs}`)
}

/** POST /spreadsheets/{id}/values:batchUpdate for several {range, values}. */
export function batchUpdateValues(client, spreadsheetId, valueInputOption, data) {
  const payload = data.map((d) => ({
    range: d.range,
    majorDimension: 'ROWS',
    values: d.values,
  }))
  return sheetsRequest(client, 'POST', `${SHEETS_API}/${spreadsheetId}/values:batchUpdate`, {
    valueInputOption,
    data: payload,
  })
}

export { SHEETS_API }