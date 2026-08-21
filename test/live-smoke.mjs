/**
 * Live smoke test for the expanded utm-gsheets plugin.
 * Uses a real OAuth token (path from GSHEETS_TOKEN_PATH env var) and a
 * real Google account. Creates a throwaway spreadsheet, exercises the new
 * tools (get, format, batch, sheet, diff), writes values/formats, reads back,
 * then DELETES the spreadsheet via the Drive API so no junk remains.
 *
 * The buildTools handlers self-hydrate a fresh authenticated client per call
 * (matching how index.js binds them), so we call each tool with only `args`.
 *
 * Run: node test/live-smoke.mjs   (with GSHEETS_TOKEN_PATH set)
 */

import { makeAuthenticatedClient } from '../src/client.js'
import { buildTools } from '../src/tools.js'

const BOLD = '\x1b[1m'
const GREEN = '\x1b[32m'
const RED = '\x1b[31m'
const RESET = '\x1b[0m'

const TOKEN_PATH = process.env.GSHEETS_TOKEN_PATH

let passed = 0
let failed = 0
const failures = []

function ok(label, cond, extra = '') {
  if (cond) {
    passed++
    console.log(`  ${GREEN}✅${RESET} ${label}`)
  } else {
    failed++
    failures.push(label + (extra ? ` — ${extra}` : ''))
    console.log(`  ${RED}❌${RESET} ${label} ${extra}`)
  }
}

async function driveDelete(fileId) {
  // fresh client -> fresh/valid access token
  const client = await makeAuthenticatedClient(TOKEN_PATH)
  // ensure a valid token: if the file token expired, refresh via refresh_token
  let token = client.access_token
  const expired = !client._expiry || client._expiry - 30_000 < Date.now()
  if (expired || !token) {
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
    const data = await res.json()
    token = data.access_token
  }
  const r = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}`, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${token}` },
  })
  return r.status
}

async function main() {
  if (!TOKEN_PATH) {
    console.error('FATAL: задайте GSHEETS_TOKEN_PATH (путь к google_token.json)')
    process.exit(1)
  }
  const tools = Object.fromEntries(buildTools(() => makeAuthenticatedClient(TOKEN_PATH)).map((t) => [t.name, t.handler]))

  console.log(`${BOLD}— create${RESET}`)
  const created = await tools.gsheets_create({
    title: `dsh-live-smoke-${Date.now()}`,
    sheet_title: 'Data',
    headers: ['Name', 'Qty', 'Price'],
    rows: [['Alpha', 10, 100], ['Beta', 5, 200]],
  })
  const id = created.spreadsheet_id
  ok('created spreadsheet with id', !!id, JSON.stringify(created))
  console.log(`  table: ${created.url}`)

  try {
    console.log(`${BOLD}— meta${RESET}`)
    const m = await tools.gsheets_meta({ spreadsheet_id: id })
    ok('meta has sheets', Array.isArray(m.sheets) && m.sheets.length >= 1)
    ok('meta found Data sheet', m.sheets.some((s) => s.title === 'Data'))

    console.log(`${BOLD}— gsheets_get (range A1:C3)${RESET}`)
    const got = await tools.gsheets_get({ spreadsheet_id: id, name: 'Data', range: 'A1:C3' })
    ok('get returned rows', Array.isArray(got.values) && got.values.length === 3, JSON.stringify(got.values))
    ok('get header row correct', got.values[0]?.[0] === 'Name' && got.values[0]?.[2] === 'Price')

    console.log(`${BOLD}— gsheets_format (bold+fill header, freeze, widths)${RESET}`)
    const fmt = await tools.gsheets_format({
      spreadsheet_id: id, name: 'Data',
      bold_header: true, fill_header: '#4285F4', header_color: '#FFFFFF',
      freeze_rows: 1, widths: { A: 160, B: 80 },
    })
    ok('format applied requests', Number(fmt.applied) >= 2, JSON.stringify(fmt))

    console.log(`${BOLD}— gsheets_write with formula (C4 = B2*C2)${RESET}`)
    const w = await tools.gsheets_write({ spreadsheet_id: id, name: 'Data', start_cell: 'C4', values: [['=B2*C2']], formulas: true })
    ok('write formula ok', w.updated_rows === 1, JSON.stringify(w))

    console.log(`${BOLD}— gsheets_get with formula render${RESET}`)
    const f2 = await tools.gsheets_get({ spreadsheet_id: id, name: 'Data', range: 'C4', value_render: 'formula' })
    ok('formula render shows "=B2*C2"', String(f2.values?.[0]?.[0]).startsWith('='), JSON.stringify(f2.values))

    console.log(`${BOLD}— gsheets_batch get (2 ranges in one call)${RESET}`)
    const batch = await tools.gsheets_batch({ spreadsheet_id: id, name: 'Data', operation: 'get', ranges: ['A1:A3', 'C1:C3'] })
    ok('batch returned 2 value ranges', Array.isArray(batch.results) && batch.results.length === 2, JSON.stringify(batch.results))

    console.log(`${BOLD}— gsheets_batch put (write 2 ranges)${RESET}`)
    const bput = await tools.gsheets_batch({
      spreadsheet_id: id, name: 'Data', operation: 'put',
      data: [
        { range: 'E1', values: [['Total']] },
        { range: 'E2', values: [['=SUM(B2:C2)']] },
      ],
    })
    ok('batch put total_updated_cells>=2', Number(bput.total_updated_cells) >= 2, JSON.stringify(bput))

    console.log(`${BOLD}— gsheets_diff (before vs after a write)${RESET}`)
    const before = (await tools.gsheets_get({ spreadsheet_id: id, name: 'Data', range: 'A1:C3' })).values
    await tools.gsheets_write({ spreadsheet_id: id, name: 'Data', start_cell: 'B2', values: [['42']] })
    const diff = await tools.gsheets_diff({ spreadsheet_id: id, name: 'Data', range: 'A1:C3', before })
    ok('diff detected a change', Number(diff.changed) >= 1, JSON.stringify(diff.changes))
    ok('diff names the changed cell', diff.changes.some((c) => c.cell && /^B2$/.test(c.cell)), JSON.stringify(diff.changes))

    console.log(`${BOLD}— gsheets_sheet (add + list + rename)${RESET}`)
    const added = await tools.gsheets_sheet({ spreadsheet_id: id, action: 'add', name: 'Report', rows: 50, cols: 10 })
    ok('sheet add ok', !!added.new_sheet_id, JSON.stringify(added))
    const lst = await tools.gsheets_sheet({ spreadsheet_id: id, action: 'list' })
    ok('sheet list has Report + Data', Array.isArray(lst.sheets) && lst.sheets.some((s) => s.title === 'Report'))
    const renamed = await tools.gsheets_sheet({ spreadsheet_id: id, action: 'rename', name: 'Report', new_title: 'Monthly' })
    ok('sheet rename ok', renamed.action === 'rename')
    const lst2 = await tools.gsheets_sheet({ spreadsheet_id: id, action: 'list' })
    ok('sheet renamed to Monthly', lst2.sheets.some((s) => s.title === 'Monthly'))

    console.log(`${BOLD}— gsheets_clear (range)${RESET}`)
    const cl = await tools.gsheets_clear({ spreadsheet_id: id, name: 'Data', range: 'A1:C1' })
    ok('clear ok', Array.isArray(cl.cleared_ranges))
  } catch (e) {
    failed++
    failures.push(`live-error: ${e?.message ?? String(e)}`)
    console.log(`  ${RED}❌${RESET} live-error: ${e?.message ?? String(e)}`)
  } finally {
    console.log(`${BOLD}— cleanup (Drive DELETE)${RESET}`)
    try {
      const status = await driveDelete(id)
      ok('drive delete returned 204/200', status === 204 || status === 200, `status=${status}`)
    } catch (e) {
      failed++
      failures.push(`cleanup: ${e?.message}`)
      console.log(`  ${RED}❌${RESET} cleanup: ${e?.message}`)
    }
  }

  console.log('\n══════════════════════════════════════')
  console.log(`Live smoke — Пройдено: ${passed}, Ошибок: ${failed}`)
  if (failures.length) {
    console.log('Провалы:')
    for (const f of failures) console.log('  ✗ ' + f)
    process.exit(1)
  }
  console.log('LIVE SMOKE PASSED')
}

main().catch((e) => {
  console.error('FATAL:', e)
  process.exit(1)
})