import type { Workbook, Worksheet } from 'exceljs'
import type { ExcelInput, ExcelResumenRow } from '../types'
import { numFmtFor } from '../format'

interface ResumenBlock {
  /** Section title — bold row above the table. */
  title: string
  /** Excel Table name (must be unique across the workbook, no spaces). */
  tableName: string
  /** Column headers — first is always "Concepto". */
  headers: string[]
  rows: ExcelResumenRow[]
  /** Number of value columns after "Concepto". */
  valueColumnCount: number
}

function maxValueColumnCount(blocks: ResumenBlock[]): number {
  return blocks.reduce((max, b) => Math.max(max, b.valueColumnCount), 0)
}

function writeBlock(sheet: Worksheet, block: ResumenBlock, startRow: number): number {
  // Title row (outside the Excel Table range).
  const titleRow = sheet.getRow(startRow)
  titleRow.getCell(1).value = block.title
  titleRow.font = { bold: true, size: 13 }

  const tableRowsStart = startRow + 1
  const tableRows = block.rows.map((row) => {
    const cells: Array<string | number | null> = [row.label]
    for (let i = 0; i < block.valueColumnCount; i += 1) {
      cells.push(row.values[i] ?? null)
    }
    return cells
  })

  if (tableRows.length === 0) {
    const headerRow = sheet.getRow(tableRowsStart)
    block.headers.forEach((label, colIdx) => {
      headerRow.getCell(colIdx + 1).value = label
    })
    headerRow.font = { bold: true }
    return tableRowsStart + 1
  }

  sheet.addTable({
    name: block.tableName,
    ref: `A${tableRowsStart}`,
    headerRow: true,
    columns: block.headers.map((name) => ({ name, filterButton: true })),
    rows: tableRows,
  })

  // Apply numFmt per data row (Resumen formats are row-level, not column-level).
  block.rows.forEach((row, rowIdx) => {
    const numFmt = numFmtFor(row.format)
    if (!numFmt) return
    const dataRow = sheet.getRow(tableRowsStart + 1 + rowIdx)
    for (let i = 0; i < block.valueColumnCount; i += 1) {
      dataRow.getCell(i + 2).numFmt = numFmt
    }
  })

  // header + data rows
  return tableRowsStart + 1 + tableRows.length
}

export function buildResumenSheet(workbook: Workbook, resumen: ExcelInput['resumen']): Worksheet {
  const sheet = workbook.addWorksheet('Resumen')
  sheet.views = [{ state: 'frozen', ySplit: 1 }]

  const blocks: ResumenBlock[] = [
    {
      title: 'Antecedentes Financieros',
      tableName: 'TableAntecedentesFinancieros',
      headers: ['Concepto', 'Titular', 'Codeudor', 'Conjunto'],
      rows: resumen.financierosRows,
      valueColumnCount: 3,
    },
    {
      title: 'Estado Situación',
      tableName: 'TableEstadoSituacion',
      headers: ['Concepto', 'Titular', 'Codeudor', 'Total'],
      rows: resumen.situacionRows,
      valueColumnCount: 3,
    },
    {
      title: 'Indicadores',
      tableName: 'TableIndicadores',
      headers: ['Concepto', 'Individual', 'Conjunto'],
      rows: resumen.indicadoresRows,
      valueColumnCount: 2,
    },
  ]

  if (resumen.edadPlazo && resumen.edadPlazo.rows.length > 0) {
    blocks.push({
      title: 'Edad + Plazo',
      tableName: 'TableEdadPlazo',
      headers: ['Concepto', ...resumen.edadPlazo.headers],
      rows: resumen.edadPlazo.rows,
      valueColumnCount: resumen.edadPlazo.headers.length,
    })
  }

  // Column widths sized for the widest block.
  const widestValueCols = maxValueColumnCount(blocks)
  sheet.columns = [
    { width: 40 },
    ...Array.from({ length: widestValueCols }, () => ({ width: 15 })),
  ]

  let cursor = 1
  blocks.forEach((block, idx) => {
    if (idx > 0) cursor += 1 // blank separator row between blocks
    cursor = writeBlock(sheet, block, cursor)
  })

  return sheet
}
