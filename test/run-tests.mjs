/**
 * Standalone tests for the utm-gsheets module. Pure Node — no DSH runtime,
 * no live Google API. Covers pure helpers (id extraction, column letters,
 * A1 ranges, parseCell, hexToRgb, request builders) and the buildTools
 * catalogue shape.
 * Run: node test/run-tests.mjs
 */

import {
  extractSpreadsheetId,
  columnLetter,
  a1Range,
  parseCell,
  parseA1Range,
  hexToRgb,
  parseSize,
  rangeObject,
  repeatCellRequest,
  basicFilterRequest,
  freezeRequest,
  dimensionSizeRequest,
  updateCellsRequest,
} from '../src/client.js'
import { buildTools } from '../src/tools.js'

let passed = 0
let failed = 0
const failures = []

function ok(label, cond, extra = '') {
  if (cond) {
    passed++
    console.log(`  ✅ ${label}`)
  } else {
    failed++
    failures.push(label + (extra ? ` — ${extra}` : ''))
    console.log(`  ❌ ${label} ${extra}`)
  }
}

function eq(label, actual, expected) {
  const a = JSON.stringify(actual)
  const e = JSON.stringify(expected)
  ok(label, a === e, `got ${a}, expected ${e}`)
}

// ---------------------------------------------------------------------------
// extractSpreadsheetId
// ---------------------------------------------------------------------------
console.log('\n— extractSpreadsheetId:')
eq('bare id', extractSpreadsheetId('1AbC123xyz'), '1AbC123xyz')
eq('full edit URL', extractSpreadsheetId('https://docs.google.com/spreadsheets/d/abc123DEF/edit#gid=0'), 'abc123DEF')
eq('share URL', extractSpreadsheetId('https://docs.google.com/spreadsheets/d/xyz789/view'), 'xyz789')
let idErr
try { extractSpreadsheetId('https://not-a-sheet/foo/bar') } catch (e) { idErr = e }
ok('invalid URL throws', idErr && /не удалось распознать/.test(idErr.message))

// ---------------------------------------------------------------------------
// columnLetter
// ---------------------------------------------------------------------------
console.log('\n— columnLetter:')
eq('col 0 -> A', columnLetter(0), 'A')
eq('col 25 -> Z', columnLetter(25), 'Z')
eq('col 26 -> AA', columnLetter(26), 'AA')
eq('col 27 -> AB', columnLetter(27), 'AB')
eq('col 51 -> AZ', columnLetter(51), 'AZ')
eq('col 52 -> BA', columnLetter(52), 'BA')
eq('col 701 -> ZZ', columnLetter(701), 'ZZ')
eq('col 702 -> AAA', columnLetter(702), 'AAA')

// ---------------------------------------------------------------------------
// a1Range
// ---------------------------------------------------------------------------
console.log('\n— a1Range:')
eq('default sheet A1 1x1', a1Range(null, 0, 0, 1, 1), 'A1:A1')
eq('default sheet A1 3x2', a1Range(null, 0, 0, 3, 2), 'A1:B3')
eq('multi-row', a1Range(null, 1, 0, 2, 3), 'A2:C3')
eq('middle offset', a1Range(null, 4, 2, 2, 2), 'C5:D6')
eq('quoted sheet name', a1Range('My Sheet', 0, 0, 1, 1), "'My Sheet'!A1:A1")
eq('quote escaping', a1Range("Bob's stuff", 0, 0, 2, 2), "'Bob''s stuff'!A1:B2")
eq('wide range', a1Range('Data', 0, 0, 50, 100), 'Data!A1:CV50')

// ---------------------------------------------------------------------------
// parseCell
// ---------------------------------------------------------------------------
console.log('\n— parseCell:')
eq('A1 -> col0 row0', parseCell('A1'), [0, 0])
eq('C5 -> col2 row4', parseCell('C5'), [2, 4])
eq('AA1 -> col26', parseCell('AA1'), [26, 0])
eq('lowercase zf10', parseCell('zf10'), [681, 9])

// ---------------------------------------------------------------------------
// parseA1Range
// ---------------------------------------------------------------------------
console.log('\n— parseA1Range:')
eq('single cell', parseA1Range('A1'), { startRow: 0, endRow: 0, startCol: 0, endCol: 0 })
eq('range', parseA1Range('A1:C3'), { startRow: 0, endRow: 2, startCol: 0, endCol: 2 })

// ---------------------------------------------------------------------------
// hexToRgb
// ---------------------------------------------------------------------------
console.log('\n— hexToRgb:')
eq('with hash', hexToRgb('#FF0000'), { red: 255, green: 0, blue: 0 })
eq('no hash', hexToRgb('00FF00'), { red: 0, green: 255, blue: 0 })
let colorErr
try { hexToRgb('#12345') } catch (e) { colorErr = e }
ok('bad hex throws', colorErr && /неверный цвет/.test(colorErr.message))

// ---------------------------------------------------------------------------
// batchUpdate request builders
// ---------------------------------------------------------------------------
console.log('\n— request builders:')
const rc = repeatCellRequest(1, 'A1:B1', { bold: true, fill: '#FF0000' })
ok('repeatCell shape', rc.repeatCell && rc.repeatCell.range.sheetId === 1)
ok('repeatCell fields includes bold', /bold/.test(rc.repeatCell.fields))
ok('repeatCell fill color applied', JSON.stringify(rc.repeatCell.cell.userEnteredFormat).includes('backgroundColor'))
ok('repeatCell range covers A1:B1', rc.repeatCell.range.startColumnIndex === 0 && rc.repeatCell.range.endColumnIndex === 2)

const bf = basicFilterRequest(3, 'A1:C5')
ok('basicFilter shape', bf.setBasicFilter.filter.range.sheetId === 3 && bf.setBasicFilter.filter.range.endRowIndex === 5)

const fz = freezeRequest(2, 1, 2)
ok('freeze shape', fz.updateSheetProperties.properties.gridProperties.frozenRowCount === 1 && fz.updateSheetProperties.properties.gridProperties.frozenColumnCount === 2)

const sz = dimensionSizeRequest(4, 'COLUMNS', 'A', '120px')
ok('col size shape', sz.updateDimensionProperties.range.sheetId === 4 && sz.updateDimensionProperties.properties.pixelSize === 120)

const ucr = updateCellsRequest(3, 'A1', [['x', 1], ['y', true]])
ok('updateCells shape', ucr.updateCells.range.sheetId === 3 && ucr.updateCells.rows.length === 2)
ok('updateCells value mapping', ucr.updateCells.rows[0][0].userEnteredValue.stringValue === 'x' && ucr.updateCells.rows[0][1].userEnteredValue.numberValue === 1)

// ---------------------------------------------------------------------------
// buildTools catalogue shape
// ---------------------------------------------------------------------------
console.log('\n— tool catalogue:')
const hydrate = () => ({ fake: true })
const tools = buildTools(hydrate)
ok('11 tools registered', tools.length === 11, `got ${tools.length}`)
const names = tools.map((t) => t.name)
eq('tool set', names, [
  'gsheets_meta', 'gsheets_list', 'gsheets_get', 'gsheets_write', 'gsheets_append',
  'gsheets_create', 'gsheets_clear', 'gsheets_sheet', 'gsheets_format', 'gsheets_batch', 'gsheets_diff',
])
ok('every tool has description+parameters', tools.every((t) => t.description && t.parameters && typeof t.handler === 'function'))
ok('spreadsheet_id required on all but create', tools.every((t) => t.name === 'gsheets_create' || t.parameters.spreadsheet_id?.required === true))

// new tools expose their key params
ok('gsheets_get has range/value_render', tools.find((t) => t.name === 'gsheets_get').parameters.range !== undefined && tools.find((t) => t.name === 'gsheets_get').parameters.value_render !== undefined)
ok('gsheets_sheet has action required', tools.find((t) => t.name === 'gsheets_sheet').parameters.action.required === true)
ok('gsheets_format has name+flags', tools.find((t) => t.name === 'gsheets_format').parameters.bold_header !== undefined && tools.find((t) => t.name === 'gsheets_format').parameters.name.required === true)
ok('gsheets_batch has operation/data', tools.find((t) => t.name === 'gsheets_batch').parameters.operation !== undefined && tools.find((t) => t.name === 'gsheets_batch').parameters.data !== undefined)
ok('gsheets_diff has before required-arg shape', tools.find((t) => t.name === 'gsheets_diff').parameters.before !== undefined)

// ---------------------------------------------------------------------------
console.log('\n══════════════════════════════════════')
console.log(`Пройдено: ${passed}, Ошибок: ${failed}`)
if (failures.length) {
  console.log('Провалы:')
  for (const f of failures) console.log('  ✗ ' + f)
  process.exit(1)
}
console.log('ALL TESTS PASSED')