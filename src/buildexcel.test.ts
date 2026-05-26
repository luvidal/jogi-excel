import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { buildExcel } from './buildexcel'
import type { ExcelApplicant, ExcelColumn, ExcelInput } from './types'

const deudasColumns: ExcelColumn[] = [
  { key: 'banco', label: 'Banco', format: 'text' },
  { key: 'tipo', label: 'Tipo', format: 'text' },
  { key: 'saldo', label: 'Saldo', format: 'currency' },
]

const propiedadesColumns: ExcelColumn[] = [
  { key: 'direccion', label: 'Dirección', format: 'text' },
  { key: 'avaluo', label: 'Avalúo', format: 'currency' },
]

const vehiculosColumns: ExcelColumn[] = [
  { key: 'marca', label: 'Marca', format: 'text' },
  { key: 'anio', label: 'Año', format: 'integer' },
]

const inversionesColumns: ExcelColumn[] = [
  { key: 'tipo', label: 'Tipo', format: 'text' },
  { key: 'monto', label: 'Monto', format: 'currency' },
]

function makeApplicant(role: 'titular' | 'codeudor', name: string, rut: string): ExcelApplicant {
  return {
    role,
    rut,
    label: name,
    edad: 40,
    perfil: [
      { section: 'Identificación', label: 'Estado civil', value: 'Casado' },
      { section: 'Laboral', subsection: 'Empleado', label: 'Empresa', value: 'Acme SA' },
    ],
    situacion: {
      deudas: {
        columns: deudasColumns,
        rows: [
          { banco: 'Banco X', tipo: 'Consumo', saldo: 1_500_000 },
          { banco: 'Banco Y', tipo: 'Hipotecario', saldo: 50_000_000 },
        ],
      },
      propiedades: {
        columns: propiedadesColumns,
        rows: [{ direccion: 'Calle 1 # 100', avaluo: 80_000_000 }],
      },
      vehiculos: {
        columns: vehiculosColumns,
        rows: [{ marca: 'Toyota', anio: 2020 }],
      },
      inversiones: {
        columns: inversionesColumns,
        rows: [{ tipo: 'Fondo mutuo', monto: 5_000_000 }],
      },
    },
  }
}

function makeEmptyApplicant(role: 'titular' | 'codeudor', name: string, rut: string): ExcelApplicant {
  return {
    role,
    rut,
    label: name,
    edad: 35,
    perfil: [{ section: 'Identificación', label: 'Estado civil', value: 'Soltero' }],
    situacion: {
      deudas: { columns: deudasColumns, rows: [] },
      propiedades: { columns: propiedadesColumns, rows: [] },
      vehiculos: { columns: vehiculosColumns, rows: [] },
      inversiones: { columns: inversionesColumns, rows: [] },
    },
  }
}

function makeInput(applicants: ExcelApplicant[]): ExcelInput {
  return {
    meta: {
      requestLabel: 'Crédito Hipotecario — Juan Pérez',
      generatedAt: '2026-05-26T12:00:00.000Z',
      ufValue: 38000,
      ufDate: '2026-05-26',
    },
    cliente: { nombre: 'Juan Pérez', rut: '11.111.111-1' },
    applicants,
    resumen: {
      financierosRows: [
        { label: 'Renta líquida', values: [3_000_000, 2_000_000, 5_000_000], format: 'currency' },
        { label: 'Dividendo', values: [800_000, 500_000, 1_300_000], format: 'currency' },
      ],
      situacionRows: [
        { label: 'Total activos', values: [85_000_000, 60_000_000, 145_000_000], format: 'currency' },
      ],
      indicadoresRows: [
        { label: 'Carga financiera', values: [0.27, 0.26], format: 'percent' },
      ],
    },
  }
}

async function loadWorkbook(input: ExcelInput): Promise<ExcelJS.Workbook> {
  const buffer = await buildExcel(input)
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer)
  return wb
}

describe('buildExcel', () => {
  it('produces 6 sheets in the correct order with the correct names', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)
    const names = wb.worksheets.map((s) => s.name)
    expect(names).toEqual(['Resumen', 'Perfil', 'Deudas', 'Propiedades', 'Vehículos', 'Inversiones'])
  })

  it('Resumen contains all three sub-tables with title rows, headers and data', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)
    const sheet = wb.getWorksheet('Resumen')!

    // Block 1: title (row 1), header (row 2), 2 data rows (rows 3-4)
    expect(sheet.getRow(1).getCell(1).value).toBe('Antecedentes Financieros')
    expect(sheet.getRow(2).getCell(1).value).toBe('Concepto')
    expect(sheet.getRow(2).getCell(2).value).toBe('Titular')
    expect(sheet.getRow(2).getCell(3).value).toBe('Codeudor')
    expect(sheet.getRow(2).getCell(4).value).toBe('Conjunto')
    expect(sheet.getRow(3).getCell(1).value).toBe('Renta líquida')
    expect(sheet.getRow(3).getCell(2).value).toBe(3_000_000)
    expect(sheet.getRow(4).getCell(1).value).toBe('Dividendo')

    // Blank row 5 → Block 2: title (row 6), header (row 7), 1 data row (row 8)
    expect(sheet.getRow(6).getCell(1).value).toBe('Estado Situación')
    expect(sheet.getRow(7).getCell(1).value).toBe('Concepto')
    expect(sheet.getRow(7).getCell(4).value).toBe('Total')
    expect(sheet.getRow(8).getCell(1).value).toBe('Total activos')
    expect(sheet.getRow(8).getCell(2).value).toBe(85_000_000)

    // Blank row 9 → Block 3: title (row 10), header (row 11), 1 data row (row 12)
    expect(sheet.getRow(10).getCell(1).value).toBe('Indicadores')
    expect(sheet.getRow(11).getCell(1).value).toBe('Concepto')
    expect(sheet.getRow(11).getCell(2).value).toBe('Individual')
    expect(sheet.getRow(11).getCell(3).value).toBe('Conjunto')
    expect(sheet.getRow(12).getCell(1).value).toBe('Carga financiera')
    expect(sheet.getRow(12).getCell(2).value).toBe(0.27)
  })

  it('Resumen wraps each block in an Excel Table with filter buttons', async () => {
    const input = makeInput([makeApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    const wb = await loadWorkbook(input)
    const sheet = wb.getWorksheet('Resumen')!
    const tableNames = (sheet as unknown as { tables: Record<string, unknown> }).tables
    expect(Object.keys(tableNames).sort()).toEqual([
      'TableAntecedentesFinancieros',
      'TableEstadoSituacion',
      'TableIndicadores',
    ])
  })

  it('Resumen renders an optional Edad+Plazo block when supplied', async () => {
    const input = makeInput([makeApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    input.resumen.edadPlazo = {
      headers: ['Titular', 'Codeudor'],
      rows: [
        { label: 'Edad + Plazo', values: [85, 90], format: 'integer' },
      ],
    }
    const wb = await loadWorkbook(input)
    const sheet = wb.getWorksheet('Resumen')!
    const tableNames = (sheet as unknown as { tables: Record<string, unknown> }).tables
    expect(Object.keys(tableNames)).toContain('TableEdadPlazo')
    // Find the Edad+Plazo title cell by scanning column 1.
    let titleRow = -1
    for (let r = 1; r <= 30; r += 1) {
      if (sheet.getRow(r).getCell(1).value === 'Edad + Plazo') { titleRow = r; break }
    }
    expect(titleRow).toBeGreaterThan(0)
    // header row immediately below title, data row below that.
    expect(sheet.getRow(titleRow + 1).getCell(1).value).toBe('Concepto')
    expect(sheet.getRow(titleRow + 1).getCell(2).value).toBe('Titular')
    expect(sheet.getRow(titleRow + 1).getCell(3).value).toBe('Codeudor')
    expect(sheet.getRow(titleRow + 2).getCell(2).value).toBe(85)
    expect(sheet.getRow(titleRow + 2).getCell(3).value).toBe(90)
  })

  it('Perfil has one row per applicant with expected base columns', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)
    const sheet = wb.getWorksheet('Perfil')!

    expect(sheet.getRow(1).getCell(1).value).toBe('Persona')
    expect(sheet.getRow(1).getCell(2).value).toBe('RUT')
    expect(sheet.getRow(1).getCell(3).value).toBe('Edad')
    expect(sheet.getRow(1).getCell(4).value).toBe('Role')
    // First perfil column header concatenates section + label.
    expect(sheet.getRow(1).getCell(5).value).toBe('Identificación — Estado civil')
    // Subsection variant.
    expect(sheet.getRow(1).getCell(6).value).toBe('Laboral — Empleado — Empresa')

    expect(sheet.getRow(2).getCell(1).value).toBe('Juan Pérez')
    expect(sheet.getRow(2).getCell(2).value).toBe('11.111.111-1')
    expect(sheet.getRow(2).getCell(3).value).toBe(40)
    expect(sheet.getRow(2).getCell(4).value).toBe('titular')
    expect(sheet.getRow(2).getCell(5).value).toBe('Casado')

    expect(sheet.getRow(3).getCell(1).value).toBe('María López')
    expect(sheet.getRow(3).getCell(4).value).toBe('codeudor')
  })

  it('situacion sheets carry the Persona + RUT prefix columns', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)

    for (const name of ['Deudas', 'Propiedades', 'Vehículos', 'Inversiones']) {
      const sheet = wb.getWorksheet(name)!
      expect(sheet.getRow(1).getCell(1).value).toBe('Persona')
      expect(sheet.getRow(1).getCell(2).value).toBe('RUT')
    }

    const deudas = wb.getWorksheet('Deudas')!
    expect(deudas.getRow(1).getCell(3).value).toBe('Banco')
    // First applicant: 2 deudas rows → rows 2-3. Second applicant: 2 deudas rows → rows 4-5.
    expect(deudas.getRow(2).getCell(1).value).toBe('Juan Pérez')
    expect(deudas.getRow(2).getCell(2).value).toBe('11.111.111-1')
    expect(deudas.getRow(2).getCell(3).value).toBe('Banco X')
    expect(deudas.getRow(2).getCell(5).value).toBe(1_500_000)
    expect(deudas.getRow(4).getCell(1).value).toBe('María López')
  })

  it('empty situacion still produces the four sheets header-only', async () => {
    const input = makeInput([makeEmptyApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    const wb = await loadWorkbook(input)

    for (const name of ['Deudas', 'Propiedades', 'Vehículos', 'Inversiones']) {
      const sheet = wb.getWorksheet(name)!
      expect(sheet.getRow(1).getCell(1).value).toBe('Persona')
      expect(sheet.getRow(1).getCell(2).value).toBe('RUT')
      // No data rows.
      expect(sheet.getRow(2).getCell(1).value).toBeFalsy()
    }
  })

  it('applies numFmt per cell on Resumen and situacion currency/percent columns', async () => {
    const input = makeInput([makeApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    const wb = await loadWorkbook(input)

    const resumen = wb.getWorksheet('Resumen')!
    // Row 3 = Renta líquida (currency, first data row of block 1, below title + header).
    expect(resumen.getRow(3).getCell(2).numFmt).toBe('#,##0')
    // Row 12 = Carga financiera (percent, first data row of block 3).
    expect(resumen.getRow(12).getCell(2).numFmt).toBe('0.00%')

    const deudas = wb.getWorksheet('Deudas')!
    // Column 5 = Saldo (currency) on row 2.
    expect(deudas.getRow(2).getCell(5).numFmt).toBe('#,##0')
  })

  it('exposes workbook metadata from input.meta', async () => {
    const input = makeInput([makeApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    const wb = await loadWorkbook(input)
    expect(wb.creator).toBe('Jogi')
    expect(wb.title).toBe('Crédito Hipotecario — Juan Pérez')
  })
})
