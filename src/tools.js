/**
 * Google Sheets tool handlers + parameter schemas.
 *
 * Thin layer over the Sheets v4 API (src/client.js). The Cordis adapter
 * (index.js) binds these through defineTool and handles error surfacing.
 *
 * Tools:
 *   gsheets_meta    — spreadsheet metadata + sheet/tab list
 *   gsheets_list    — read cells (records or raw arrays); optional A1 range
 *   gsheets_get     — read one exact A1 range (+ formula / unformatted render)
 *   gsheets_write   — overwrite a rectangle starting at a cell
 *   gsheets_append  — append rows to the end of a sheet
 *   gsheets_create  — create a new spreadsheet (+ optional sheet/data)
 *   gsheets_clear   — clear all/part of a sheet (values only)
 *   gsheets_sheet   — add / rename / delete / resize / duplicate a tab
 *   gsheets_format  — formatting: header bold/fill, autofilter, freeze, widths/heights
 *   gsheets_batch   — batch read (get) or batch write (put) ranges at once
 *   gsheets_diff    — compare a before snapshot with current values
 */

import {
  SHEETS_API,
  extractSpreadsheetId,
  a1Range,
  columnLetter,
  parseCell,
  repeatCellRequest,
  basicFilterRequest,
  freezeRequest,
  dimensionSizeRequest,
  batchUpdate,
  batchGetValues,
  batchUpdateValues,
  sheetsRequest,
} from './client.js'

// ---------------------------------------------------------------------------
// URL / range helpers
// ---------------------------------------------------------------------------

const SPREADSHEET_URL = (id) => `${SHEETS_API}/${id}`
const VALUES_URL = (id) => `${SPREADSHEET_URL(id)}/values`

/** Quote a sheet title for an embedded range if it contains spaces/specials. */
export function quoteSheet(title) {
  return /\s/.test(String(title)) ? `'${String(title).replace(/'/g, "''")}'` : String(title)
}

/** Full range string: optional quoted sheet + A1. */
function sheetRange(sheetName, range) {
  const body = range ? String(range) : 'A1'
  return sheetName ? `${quoteSheet(sheetName)}!${body}` : body
}

/** Fetch spreadsheet metadata (sheets list + grid sizes). */
async function fetchMeta(client, id) {
  return sheetsRequest(client, 'GET', `${SPREADSHEET_URL(id)}?includeGridData=false`)
}

/** Resolve a sheet title to its metadata entry (case-insensitive). */
function findSheet(meta, sheetName) {
  const want = String(sheetName ?? '').trim().toLowerCase()
  if (!want) return meta?.sheets?.[0] ?? null
  return (meta?.sheets ?? []).find(
    (s) => String(s.properties?.title ?? '').toLowerCase() === want,
  ) ?? null
}

/** Shape a sheet metadata entry for output. */
function sheetMeta(s) {
  return {
    title: s.properties?.title ?? '',
    sheet_id: s.properties?.sheetId ?? null,
    index: s.properties?.index ?? null,
    row_count: s.properties?.gridProperties?.rowCount ?? 0,
    col_count: s.properties?.gridProperties?.columnCount ?? 0,
  }
}

/** Normalize a raw values page: pad each row to the widest row length. */
function normalizeGrid(rows) {
  const raw = rows ?? []
  const max = raw.reduce((m, r) => Math.max(m, (r ?? []).length), 0)
  return raw.map((r, i) => Array.from({ length: max }, (_, c) => (r ?? [])[c] ?? ''))
}

/** Convert a values grid + header row index to record objects. */
function gridToRecords(values, headerRowInx, keepHeader) {
  const grid = (values ?? []).map((r) => [...(r ?? [])])
  if (!grid.length) return []
  const header = headerRowInx === 0 && !keepHeader && grid.length ? grid[0] : grid[headerRowInx] ?? []
  const startRow = headerRowInx === 0 && !keepHeader ? 1 : headerRowInx + (keepHeader ? 0 : 1)
  const cols = Math.max(header.length, ...grid.slice(startRow).map((r) => r.length).concat([0]))
  const records = []
  for (let i = Math.min(startRow, grid.length); i < grid.length; i++) {
    const row = grid[i] ?? []
    const rec = {}
    for (let c = 0; c < cols; c++) {
      rec[header[c] != null && String(header[c]).trim() ? String(header[c]) : `col${c + 1}`] =
        row[c] == null ? '' : row[c]
    }
    if (Object.values(rec).every((v) => v === '')) continue
    records.push(rec)
  }
  return records
}

// ---------------------------------------------------------------------------
// Handlers
// ---------------------------------------------------------------------------

async function meta(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const data = await fetchMeta(client, id)
  return {
    spreadsheet_id: id,
    title: data.properties?.title ?? '',
    locale: data.properties?.locale ?? null,
    timezone: data.properties?.timeZone ?? null,
    url: `https://docs.google.com/spreadsheets/d/${id}/edit`,
    sheets: (data.sheets ?? []).map(sheetMeta),
  }
}

/** Normalize a user `values` argument into a 2D grid. */
function normalizeInput(values) {
  const grid = Array.isArray(values) ? values : []
  if (!grid.length) throw new Error('values: нужен непустой массив строк (массив массивов)')
  return Array.isArray(grid[0]) ? grid : [grid]
}

async function getRange(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const sheet = args.sheet ?? null
  const range = args.range ?? 'A1'
  const opt = args.value_render ?? 'formatted'
  const render =
    opt === 'formula' ? 'FORMULA' : opt === 'unformatted' ? 'UNFORMATTED_VALUE' : 'FORMATTED_VALUE'
  const full = sheetRange(sheet, range)
  const data = await sheetsRequest(client, 'GET', `${VALUES_URL(id)}/${encodeURIComponent(full)}?valueRenderOption=${render}&majorDimension=ROWS`)
  const values = normalizeGrid(data.values)
  return { spreadsheet_id: id, sheet, range: full, rows: values.length, values }
}

async function listRows(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const headerRow = Number(args.header_row ?? 0)
  const limit = Number(args.limit ?? 200)
  const asArrays = Boolean(args.as_arrays)
  const keepHeader = Boolean(args.keep_header)
  const sheet = args.sheet ?? null
  const range = args.range ?? 'A1:ZZZ10000'
  const full = sheetRange(sheet, range)
  const data = await sheetsRequest(client, 'GET', `${VALUES_URL(id)}/${encodeURIComponent(full)}?valueRenderOption=FORMATTED_VALUE&majorDimension=ROWS`)
  const raw = data.values ?? []
  if (asArrays) {
    const skip = headerRow + (!keepHeader && headerRow === 0 ? 1 : 0)
    const values = raw.slice(skip, skip + limit)
    return {
      spreadsheet_id: id, sheet, header_row: headerRow,
      rows_total: raw.length, rows_returned: values.length,
      truncated: raw.slice(skip).length > limit,
      values,
    }
  }
  const records = gridToRecords(raw, headerRow, keepHeader)
  return {
    spreadsheet_id: id, sheet, header_row: headerRow,
    rows_total: raw.length, rows_returned: records.length,
    truncated: records.length > limit,
    items: records.slice(0, limit),
  }
}

async function writeCells(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const grid = normalizeInput(args.values)
  if (!grid.every((r) => Array.isArray(r))) throw new Error('values: каждая строка — массив значений')
  const [startCol, startRow] = parseCell(args.start_cell ?? 'A1')
  const nCols = Math.max(...grid.map((r) => r.length))
  const range = a1Range(args.sheet ?? null, startRow, startCol, grid.length, nCols)
  const opt = args.user_input || args.formulas ? 'USER_ENTERED' : 'RAW'
  const data = await sheetsRequest(client, 'PUT', `${VALUES_URL(id)}/${encodeURIComponent(range)}?valueInputOption=${opt}`, { majorDimension: 'ROWS', values: grid })
  return {
    spreadsheet_id: id,
    range: data.updatedRange ?? range,
    updated_cells: data.updatedCells ?? grid.length * nCols,
    updated_rows: grid.length,
    updated_columns: nCols,
  }
}

async function appendRows(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const grid = normalizeInput(args.values)
  if (!grid.every((r) => Array.isArray(r))) throw new Error('values: каждая строка — массив значений')
  const opt = args.user_input || args.formulas ? 'USER_ENTERED' : 'RAW'
  const url = `${VALUES_URL(id)}/${encodeURIComponent(sheetRange(args.sheet ?? '', 'A1'))}:append?valueInputOption=${opt}&insertDataOption=INSERT_ROWS`
  const data = await sheetsRequest(client, 'POST', url, { majorDimension: 'ROWS', values: grid })
  return {
    spreadsheet_id: id,
    table_range: data.tableRange ?? null,
    updated_range: data.updates?.updatedRange ?? null,
    updated_rows: data.updates?.updatedRows ?? grid.length,
    updated_cells: data.updates?.updatedCells ?? null,
  }
}

async function createSpreadsheet(client, args) {
  const title = String(args.title ?? 'Untitled')
  const headers = Array.isArray(args.headers) ? args.headers : []
  const rows = Array.isArray(args.rows) ? args.rows : []
  const firstSheet = args.sheet_title ? String(args.sheet_title) : null

  const payload = { properties: { title } }
  if (firstSheet) {
    payload.sheets = [
      { properties: { title: firstSheet, gridProperties: { rowCount: 1000, columnCount: 100 } } },
    ]
  }
  const data = await sheetsRequest(client, 'POST', SPREADSHEET_URL(''), payload)
  const id = data.spreadsheetId
  const sheetName = firstSheet ?? data.sheets?.[0]?.properties?.title ?? null

  let populated = null
  if (headers.length || rows.length) {
    const grid = [headers.length ? headers : null, ...rows].filter((g) => Array.isArray(g) && g.length)
    if (grid.length) {
      const range = a1Range(sheetName, 0, 0, grid.length, Math.max(...grid.map((r) => r.length)))
      const w = await sheetsRequest(client, 'PUT', `${VALUES_URL(id)}/${encodeURIComponent(range)}?valueInputOption=RAW`, { values: grid })
      populated = { range: w.updatedRange ?? range, rows: grid.length }
    }
  }

  return {
    spreadsheet_id: id,
    url: `https://docs.google.com/spreadsheets/d/${id}/edit`,
    title,
    sheet: sheetName,
    populated,
  }
}

async function clearCells(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const sheet = args.sheet ?? null
  const rangeSpec = args.range ?? null
  const ranges = [
    rangeSpec
      ? sheet ? `${quoteSheet(sheet)}!${rangeSpec}` : rangeSpec
      : sheet ? `${quoteSheet(sheet)}!A1:ZZZ10000` : 'A1:ZZZ10000',
  ]
  const data = await sheetsRequest(client, 'POST', `${SPREADSHEET_URL(id)}/values:batchClear`, { ranges })
  return { spreadsheet_id: id, cleared_ranges: data.clearedRanges ?? ranges }
}

async function sheetOps(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const action = String(args.action ?? 'list')
  const metaData = await fetchMeta(client, id)

  if (action === 'list') {
    return { spreadsheet_id: id, sheets: (metaData.sheets ?? []).map(sheetMeta) }
  }

  if (action === 'duplicate') {
    const target = findSheet(metaData, args.name)
    if (!target) throw new Error(`лист не найден: ${args.name}`)
    const r = await sheetsRequest(client, 'POST', `${SPREADSHEET_URL(id)}:sheets/duplicateRequest`, {
      sourceSheetId: target.properties.sheetId,
    })
    const newId = r?.replies?.[0]?.duplicateSheet?.newSheetId ?? null
    let newTitle = null
    if (args.name && args.new_title) {
      await batchUpdate(client, id, [
        { updateSheetProperties: { properties: { sheetId: newId, title: String(args.new_title) }, fields: 'title' } },
      ])
      newTitle = String(args.new_title)
    }
    return { spreadsheet_id: id, action, new_sheet_id: newId, new_title: newTitle ?? args.name }
  }

  let requests = []
  let notes = []
  let newSheetId = null
  let newTitle = null

  switch (action) {
    case 'add': {
      const title = String(args.name || args.new_title || `Лист ${(metaData.sheets?.length ?? 0) + 1}`)
      const rowsN = Number(args.rows) || 1000
      const colsN = Number(args.cols) || 100
      requests.push({
        addSheet: {
          properties: { title, gridProperties: { rowCount: rowsN, columnCount: colsN } },
        },
      })
      newTitle = title
      break
    }
    case 'rename': {
      const target = findSheet(metaData, args.name)
      if (!target) throw new Error(`лист не найден: ${args.name}`)
      const newName = String(args.new_title ?? '').trim()
      if (!newName) throw new Error('передайте new_title для rename')
      requests.push({
        updateSheetProperties: { properties: { sheetId: target.properties.sheetId, title: newName }, fields: 'title' },
      })
      notes.push(`переименован "${target.properties.title}" -> "${newName}"`)
      newTitle = newName
      break
    }
    case 'delete': {
      const target = findSheet(metaData, args.name)
      if (!target) throw new Error(`лист не найден: ${args.name}`)
      requests.push({ deleteSheet: { sheetId: target.properties.sheetId } })
      notes.push(`удалён лист «${target.properties.title}»`)
      break
    }
    case 'resize': {
      const target = findSheet(metaData, args.name)
      if (!target) throw new Error(`лист не найден: ${args.name}`)
      const g = target.properties.gridProperties ?? {}
      requests.push({
        updateSheetProperties: {
          properties: {
            sheetId: target.properties.sheetId,
            gridProperties: {
              rowCount: Number(args.rows) || g.rowCount || 1000,
              columnCount: Number(args.cols) || g.columnCount || 100,
            },
          },
          fields: 'gridProperties.rowCount,gridProperties.columnCount',
        },
      })
      notes.push(`изменён размер листа «${target.properties.title}»`)
      break
    }
    default:
      throw new Error(`gsheets_sheet: неизвестное действие ${action} (add|rename|delete|resize|duplicate|list)`)
  }

  const out = await batchUpdate(client, id, requests)
  const add = out?.replies?.[0]?.addSheet
  if (add) {
    newSheetId = add.properties?.sheetId ?? null
    newTitle = add.properties?.title ?? null
  }
  return {
    spreadsheet_id: id,
    action,
    new_sheet_id: newSheetId,
    new_title: newTitle,
    notes,
  }
}

/** Formatting via batchUpdate against one sheet. */
async function formatCells(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const metaData = await fetchMeta(client, id)
  const target = findSheet(metaData, args.name)
  if (!target) throw new Error(`лист не найден: ${args.name}`)
  const sid = target.properties.sheetId
  const g = target.properties.gridProperties ?? {}
  const rows = Number(args.rows) || g.rowCount || 1000
  const cols = Number(args.cols) || g.columnCount || 100

  const requests = []
  if (args.bold_header != null || args.fill_header || args.header_color) {
    const fmt = {}
    if (args.bold_header != null) fmt.bold = Boolean(args.bold_header)
    if (args.header_color) fmt.color = args.header_color
    if (args.fill_header) fmt.fill = args.fill_header
    requests.push(repeatCellRequest(sid, `A1:${columnLetter(cols - 1)}1`, fmt))
  }
  if (args.autofilter) {
    requests.push(basicFilterRequest(sid, `A1:${columnLetter(cols - 1)}${rows}`))
  }
  if (args.freeze_rows || args.freeze_cols) {
    requests.push(freezeRequest(sid, Number(args.freeze_rows) || 0, Number(args.freeze_cols) || 0))
  }
  if (args.widths && typeof args.widths === 'object') {
    for (const [lett, val] of Object.entries(args.widths)) {
      if (!/^[A-Z]+$/i.test(lett)) continue
      requests.push(dimensionSizeRequest(sid, 'COLUMNS', String(lett).toUpperCase(), String(val)))
    }
  }
  if (args.heights && typeof args.heights === 'object') {
    for (const [rn, val] of Object.entries(args.heights)) {
      if (!/^\d+$/.test(rn)) continue
      requests.push(dimensionSizeRequest(sid, 'ROWS', rn, String(val)))
    }
  }
  if (!requests.length) {
    throw new Error('не передано ни одного действия форматирования (bold_header, fill_header, autofilter, freeze_*, widths, heights)')
  }

  const out = await batchUpdate(client, id, requests)
  return {
    spreadsheet_id: id,
    sheet: target.properties.title ?? '',
    applied: out?.replies?.length ?? 0,
  }
}

async function batchOp(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const action = String(args.operation ?? 'get')

  if (action === 'get') {
    const list = Array.isArray(args.ranges) ? args.ranges : args.ranges ? [args.ranges] : []
    if (!list.length) throw new Error('ranges: передайте массив A1-диапазонов')
    const ranges = list.map((r) =>
      typeof r === 'string'
        ? sheetRange(args.name ?? null, r)
        : sheetRange(r?.name ?? args.name ?? null, r?.range ?? 'A1'))
    const data = await batchGetValues(client, id, ranges)
    return {
      spreadsheet_id: id,
      ranges: data,
      results: (data.valueRanges ?? []).map((vr) => ({
        range: vr.range ?? null,
        values: normalizeGrid(vr.values),
      })),
    }
  }

  if (action === 'put') {
    const list = Array.isArray(args.data) ? args.data : args.data ? [args.data] : []
    if (!list.length) throw new Error('data: передайте массив записей {range, values}')
    const payload = list.map((d) => ({
      range: sheetRange(d.name ?? args.name ?? null, d.range ?? d.name ?? 'A1'),
      values: normalizeInput(d.values),
    }))
    const out = await batchUpdateValues(client, id, args.value_input ? 'USER_ENTERED' : 'RAW', payload)
    return {
      spreadsheet_id: id,
      total_updated_cells: out.totalUpdatedCells ?? null,
      responses: (out.responses ?? []).map((r) => ({
        range: r.updatedRange ?? null,
        updated_cells: r.updatedCells ?? null,
      })),
    }
  }
  throw new Error(`gsheets_batch: неизвестное действие ${action} (get|put)`)
}

/** Compare `before` snapshot with current values. */
async function diffCells(client, args) {
  const id = extractSpreadsheetId(args.spreadsheet_id)
  const sheet = args.name ?? null
  const range = args.range ?? 'A1:ZZZ10000'
  if (!Array.isArray(args.before)) throw new Error('before: передайте массив строк значений ДО')
  const opt = args.value_render ?? 'formatted'
  const render = opt === 'formula' ? 'FORMULA' : 'FORMATTED_VALUE'
  const full = sheetRange(sheet, range)
  const curData = await sheetsRequest(client, 'GET', `${VALUES_URL(id)}/${encodeURIComponent(full)}?valueRenderOption=${render}&majorDimension=ROWS`)
  const before = normalizeGrid(args.before)
  const after = normalizeGrid(curData.values)
  const rows = Math.max(before.length, after.length)
  const cols = Math.max(1, ...before.map((r) => r.length), ...after.map((r) => r.length))
  const changes = []
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const oldV = before[r]?.[c] ?? ''
      const newV = after[r]?.[c] ?? ''
      if (String(oldV) !== String(newV)) {
        changes.push({ cell: `${columnLetter(c)}${r + 1}`, old: oldV, new: newV })
      }
    }
  }
  return { spreadsheet_id: id, range: full, changed: changes.length, changes }
}

// ---------------------------------------------------------------------------
// Catalogue — hydrate() builds a fresh authenticated client per call
// ---------------------------------------------------------------------------

export function buildTools(hydrate) {
  return [
    {
      name: 'gsheets_meta',
      description: 'Метаданные Google-таблицы: заголовок, locale, timezone и список листов (tabs) с числом строк/колонок. Принимает URL или id. Read-only.',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'ID таблицы или полный URL (https://docs.google.com/spreadsheets/d/<ID>/edit).' },
      },
      handler: async (args) => meta(await hydrate(), args),
    },
    {
      name: 'gsheets_list',
      description: 'Прочитать данные листа Google-таблицы. range — необязательный точный A1-диапазон (иначе весь лист). as_arrays=true — сырые строки; keep_header=true — не отбрасывать строку-заголовок. Read-only.',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        name: { type: 'string', description: 'Название листа (tab). Пусто — первый лист.' },
        range: { type: 'string', description: 'Точный A1-диапазон, например "A1:C50". Пусто — весь лист.' },
        header_row: { type: 'integer', description: 'Индекс строки-заголовка (0-based). По умолчанию 0.' },
        limit: { type: 'integer', description: 'Максимум строк в ответе (по умолчанию 200).' },
        as_arrays: { type: 'boolean', description: 'True — вернуть сырые строки (массив массивов).' },
        keep_header: { type: 'boolean', description: 'При header_row=0 не отбрасывать первую строку.' },
      },
      handler: async (args) => listRows(await hydrate(), args),
    },
    {
      name: 'gsheets_get',
      description: 'Прочитать точный A1-диапазон листа как сырые строки. value_render: formatted (как в UI) | formula (сами формулы) | unformatted (сырые числа). Read-only.',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        name: { type: 'string', description: 'Название листа (tab). Пусто — первый.' },
        range: { type: 'string', description: 'A1-диапазон, например "A1:C50". По умолчанию A1.' },
        value_render: { type: 'string', description: 'formatted | formula | unformatted (по умолчанию formatted).' },
      },
      handler: async (args) => getRange(await hydrate(), args),
    },
    {
      name: 'gsheets_write',
      description: 'Записать прямоугольник значений, начиная с указанной ячейки. values — массив строк (массив массивов). RAW по умолчанию; user_input=true (или formulas=true) — разбор как в Google Sheets: даты, проценты, формулы (=A1+B1).',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        values: { type: 'array', required: true, description: 'Массив строк (массив массивов) или одна строка.', items: { type: 'array', items: { oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }] } } },
        name: { type: 'string', description: 'Лист (tab). Пусто — первый.' },
        start_cell: { type: 'string', description: 'Начальная ячейка, например A1/C3. По умолчанию A1.' },
        user_input: { type: 'boolean', description: 'True — разобрать ввод как в Google Sheets.' },
        formulas: { type: 'boolean', description: 'True — писать формулы (=…). То же user_input.' },
      },
      handler: async (args) => writeCells(await hydrate(), args),
    },
    {
      name: 'gsheets_append',
      description: 'Добавить строки в конец листа (после последней строки с данными). values — массив строк (ROW-wise).',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        values: { type: 'array', required: true, description: 'Строки для добавления, например [[1,"a"],[2,"b"]].', items: { type: 'array', items: { oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }] } } },
        name: { type: 'string', description: 'Название листа. Пусто — первый.' },
        user_input: { type: 'boolean', description: 'True — разобрать ввод как в Google Sheets.' },
      },
      handler: async (args) => appendRows(await hydrate(), args),
    },
    {
      name: 'gsheets_create',
      description: 'Создать новую Google-таблицу: title, опционально первый лист (sheet_title), headers и строки данных (rows). Возвращает id и URL.',
      parameters: {
        title: { type: 'string', description: 'Название таблицы.' },
        headers: { type: 'array', items: { type: 'string' }, description: 'Заголовки колонок (строка 1).' },
        rows: { type: 'array', items: { type: 'array', items: { oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }] } }, description: 'Данные строк, пишутся с row 2.' },
        sheet_title: { type: 'string', description: 'Название первого листа.' },
      },
      handler: async (args) => createSpreadsheet(await hydrate(), args),
    },
    {
      name: 'gsheets_clear',
      description: 'Очистить значения листа (или диапазона). Форматирование и структуру не трогает.',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        name: { type: 'string', description: 'Название листа. Без range — весь лист.' },
        range: { type: 'string', description: 'A1-диапазон для очистки, например A1:C10.' },
      },
      handler: async (args) => clearCells(await hydrate(), args),
    },
    {
      name: 'gsheets_sheet',
      description: 'Управление вкладками таблицы. Действия: add (создать), rename, delete, resize, duplicate, list.',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        action: { type: 'string', required: true, description: 'add | rename | delete | resize | duplicate | list' },
        name: { type: 'string', description: 'Название существующего листа (rename/delete/resize/duplicate ду и для add через new_title).' },
        new_title: { type: 'string', description: 'Название нового листа (add) или новое название (rename/duplicate).' },
        rows: { type: 'integer', description: 'Для add/resize — число строк.' },
        cols: { type: 'integer', description: 'Для add/resize — число колонок.' },
      },
      handler: async (args) => sheetOps(await hydrate(), args),
    },
    {
      name: 'gsheets_format',
      description: 'Форматирование листа через batchUpdate: bold_header, fill_header (#RRGGBB), header_color, autofilter, freeze_rows, freeze_cols, widths {"A":150}, heights {"1":40}.',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        name: { type: 'string', required: true, description: 'Название листа (tab).' },
        bold_header: { type: 'boolean', description: 'Сделать первый ряд жирным.' },
        fill_header: { type: 'string', description: 'Цвет фона шапки, #RRGGBB.' },
        header_color: { type: 'string', description: 'Цвет текста шапки, #RRGGBB.' },
        autofilter: { type: 'boolean', description: 'Включить фильтр по всей области.' },
        freeze_rows: { type: 'integer', description: 'Закрепить N верхних строк (0 — отменить).' },
        freeze_cols: { type: 'integer', description: 'Закрепить N левых колонок.' },
        widths: { type: 'object', description: 'Ширины колонок: {"A":150,"B":120}.' },
        heights: { type: 'object', description: 'Высоты строк: {"1":40}.' },
      },
      handler: async (args) => formatCells(await hydrate(), args),
    },
    {
      name: 'gsheets_batch',
      description: 'Пакетная операция. operation=get — прочитать несколько диапазонов одним запросом; operation=put — записать несколько диапазонов.',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        operation: { type: 'string', description: 'get | put (по умолчанию get).' },
        name: { type: 'string', description: 'Целевой лист по умолчанию.' },
        ranges: { type: 'array', items: { oneOf: [{ type: 'string' }, { type: 'object' }] }, description: 'Для get: A1-диапазоны (строки или {range}).' },
        data: { type: 'array', items: { type: 'object' }, description: 'Для put: записи {range, values}.' },
        value_input: { type: 'boolean', description: 'Для put — разобрать как Google Sheets.' },
      },
      handler: async (args) => batchOp(await hydrate(), args),
    },
    {
      name: 'gsheets_diff',
      description: 'Сравнить переданный before (массив строк) с текущими значениями листа и вернуть изменённые ячейки (старое -> новое).',
      parameters: {
        spreadsheet_id: { type: 'string', required: true, description: 'URL или id таблицы.' },
        name: { type: 'string', description: 'Название листа.' },
        range: { type: 'string', description: 'A1-диапазон для сравнения. По умолчанию весь лист.' },
        before: { type: 'array', description: 'До-снимок значений (массив строк).', items: { type: 'array', items: { oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }] } } },
        value_render: { type: 'string', description: 'formatted | formula (чем сравнивать).' },
      },
      handler: async (args) => diffCells(await hydrate(), args),
    },
  ]
}