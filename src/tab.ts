import type { sheets_v4 } from "@googleapis/sheets"
import { ZodError, type z } from "zod"
import { Cell } from "./cell"
import { Row } from "./row"
import {
  ValidationError,
  type SheetInstance,
  type ValueInputOption,
  type ValueRenderOption,
} from "./types"
import {
  columnLabel,
  columnIndex,
  parseCellAddress,
  extractCellValue,
  extractCellFormat,
  validateArrayAgainstSchema,
} from "./utils"

export interface AppendOpts {
  values: Record<string, unknown> | unknown[]
  inputFormat?: ValueInputOption
  getRow?: boolean
}

export interface ValuesOpts {
  range?: string
  valueRenderOption?: ValueRenderOption
}

export interface GetRowOpts {
  serialNo: number
  valueRenderOption?: ValueRenderOption
}

export interface GetCellOpts {
  cellName: string
  valueRenderOption?: ValueRenderOption
}

export interface DelRowOpts {
  start: number
  end?: number
}

export interface DelCellOpts {
  start: string
  end?: string
  shift?: "up" | "left"
}

export class Tab {
  private sheet: SheetInstance
  private title: string
  private worksheetId: number
  private _gridProperties: sheets_v4.Schema$GridProperties
  private _headers: string[] = []
  private _schema: z.ZodObject | null = null

  constructor(
    sheet: SheetInstance,
    title: string,
    worksheetId: number,
    gridProperties: sheets_v4.Schema$GridProperties = {},
  ) {
    this.sheet = sheet
    this.title = title
    this.worksheetId = worksheetId
    this._gridProperties = gridProperties
  }

  getHeaders(): string[] {
    return this._headers
  }

  _setHeaders(headers: string[]): void {
    // Always preserve real sheet header order
    this._headers = headers

    // If schema is set, validate all schema keys exist in actual headers
    if (this._schema) {
      const schemaKeys = Object.keys(this._schema.shape)
      const missingKeys = schemaKeys.filter((key) => !headers.includes(key))
      if (missingKeys.length > 0) {
        throw new Error(
          `Schema keys not found in sheet headers: ${missingKeys.join(", ")}. ` +
            `Sheet headers: ${headers.join(", ")}`,
        )
      }
    }
  }

  setSchema(schema: z.ZodObject): this {
    this._schema = schema

    // If headers already cached, validate schema keys match
    if (this._headers.length > 0) {
      const schemaKeys = Object.keys(schema.shape)
      const missingKeys = schemaKeys.filter(
        (key) => !this._headers.includes(key),
      )
      if (missingKeys.length > 0) {
        throw new Error(
          `Schema keys not found in sheet headers: ${missingKeys.join(", ")}. ` +
            `Sheet headers: ${this._headers.join(", ")}`,
        )
      }
    }

    return this
  }

  getSchema(): z.ZodObject | null {
    return this._schema
  }

  getClient(): sheets_v4.Sheets {
    return this.sheet.getClient()
  }

  getSheetId(): string {
    return this.sheet.getId()
  }

  getTitle(): string {
    return this.title
  }

  getWorksheetId(): number {
    return this.worksheetId
  }

  async values(opts?: ValuesOpts): Promise<Row[]> {
    const client = this.getClient()
    const { range, valueRenderOption } = opts ?? {}
    const actualRange = range ?? `${this.title}!A1:ZZZ`

    // When render option specified, use spreadsheets.values.get (no format data)
    if (valueRenderOption) {
      const res = await client.spreadsheets.values.get({
        spreadsheetId: this.getSheetId(),
        range: actualRange,
        valueRenderOption,
      })
      const rows = (res.data.values as unknown[][]) ?? []

      // Cache headers from first row
      if (rows[0]) {
        this._setHeaders(rows[0].map((v) => String(v ?? "")))
      }

      return rows.map((row, i) => {
        const cells = row.map((val, j) => {
          return new Cell({
            value: String(val ?? ""),
            label: columnLabel(j),
            rowIndex: i + 1,
            cellIndex: j,
            tab: this,
            row: null,
          })
        })
        return new Row(cells, this, i + 1)
      })
    }

    // Default path: spreadsheets.get with effectiveValue + cell format
    const res = await client.spreadsheets.get({
      spreadsheetId: this.getSheetId(),
      ranges: [actualRange],
      fields: "sheets.data.rowData.values(effectiveValue,userEnteredFormat)",
    })

    const rowData = res.data.sheets?.[0]?.data?.[0]?.rowData ?? []

    // Cache headers from first row
    const firstRow = rowData[0]
    if (firstRow?.values) {
      this._setHeaders(
        firstRow.values.map((v) =>
          String(extractEffectiveValue(v.effectiveValue) ?? ""),
        ),
      )
    }

    return rowData.map((row, i) => {
      const values = row.values ?? []
      const cells = values.map((v, j) => {
        const rawValue = extractEffectiveValue(v.effectiveValue)
        const cellFormat = v.userEnteredFormat ?? null
        return new Cell({
          value: String(rawValue ?? ""),
          label: columnLabel(j),
          rowIndex: i + 1,
          cellIndex: j,
          tab: this,
          row: null,
          format: cellFormat ?? undefined,
        })
      })
      return new Row(cells, this, i + 1)
    })
  }

  async getRow(opts: GetRowOpts): Promise<Row> {
    const { serialNo, valueRenderOption } = opts
    const rows = await this.values({
      range: `${this.title}!A${serialNo}:ZZZ${serialNo}`,
      valueRenderOption,
    })
    return rows[0] ?? new Row([], this, serialNo)
  }

  async getCell(opts: GetCellOpts): Promise<Cell> {
    const { cellName, valueRenderOption = "FORMATTED_VALUE" } = opts
    const client = this.getClient()

    const res = await client.spreadsheets.values.get({
      spreadsheetId: this.getSheetId(),
      range: `${this.title}!${cellName}`,
      valueRenderOption,
    })

    const value = (res.data.values?.[0]?.[0] as string) ?? ""
    const addr = parseCellAddress(cellName)
    const label = addr?.label ?? "A"
    const cellIndex = columnIndex(label)
    const rowIndex =
      addr?.row ?? parseInt(cellName.match(/\d+/)?.[0] ?? "1", 10)

    return new Cell({
      value,
      label,
      rowIndex,
      cellIndex,
      tab: this,
      row: null,
    })
  }

  private objectToRow(data: Record<string, unknown>): unknown[] {
    if (this._headers.length === 0) {
      throw new Error(
        "Tab headers not cached. Call values() first or set headers manually.",
      )
    }
    if (this._schema) {
      try {
        this._schema.parse(data)
      } catch (err) {
        if (err instanceof ZodError) {
          throw new ValidationError(
            `Schema validation failed: ${err.message}`,
            err,
          )
        }
        throw err
      }
    }
    return this._headers.map((h) => {
      const val = data[h]
      if (val === undefined || val === null) return ""
      if (val instanceof Cell) return val.value
      return String(val)
    })
  }

  async append(opts: AppendOpts): Promise<Row | null> {
    const client = this.getClient()
    const { values: data, inputFormat, getRow } = opts

    const rowArray = Array.isArray(data) ? data : this.objectToRow(data)

    // Validate array against schema by column position
    if (this._schema) {
      if (this._headers.length === 0) {
        throw new Error(
          "Tab headers not cached. Call values() first or use object append.",
        )
      }
      validateArrayAgainstSchema(rowArray, this._headers, this._schema)
    }

    // Write values via values.append — the API echoes the exact landing range
    // (updates.updatedRange), so the appended row index is known without a
    // full-sheet read and without racing concurrent appends. appendCells gave
    // no such echo, and gridProperties.rowCount is the grid size (often 1000),
    // not the last data row, so it could not be used to locate the new row.
    const appendRes = await client.spreadsheets.values.append({
      spreadsheetId: this.getSheetId(),
      range: `${this.title}!A1`,
      valueInputOption: inputFormat ?? "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [rowArray.map((v) => extractCellValue(v))],
      },
    })

    const updatedRange = appendRes.data.updates?.updatedRange
    const appendedRowIndex = updatedRange
      ? parseAppendedRowIndex(updatedRange)
      : null

    // values.append cannot carry per-cell format. Apply formats in a second
    // pass, only when at least one input cell actually has a format.
    const cellFormats = rowArray.map((v) => extractCellFormat(v))
    const hasFormat = cellFormats.some((f) => f !== undefined)
    if (hasFormat && appendedRowIndex != null) {
      await client.spreadsheets.batchUpdate({
        spreadsheetId: this.getSheetId(),
        requestBody: {
          requests: [
            {
              updateCells: {
                range: {
                  sheetId: this.worksheetId,
                  startRowIndex: appendedRowIndex - 1,
                  endRowIndex: appendedRowIndex,
                  startColumnIndex: 0,
                  endColumnIndex: cellFormats.length,
                },
                rows: [
                  {
                    values: cellFormats.map((f) =>
                      f ? { userEnteredFormat: f } : {},
                    ),
                  },
                ],
                fields: "userEnteredFormat",
              },
            },
          ],
        },
      })
    }

    if (!getRow || appendedRowIndex == null) return null

    // Read the landed row (effectiveValue + applied format) at the echoed index.
    const res = await client.spreadsheets.get({
      spreadsheetId: this.getSheetId(),
      ranges: [`${this.title}!A${appendedRowIndex}:ZZZ${appendedRowIndex}`],
      fields: "sheets.data.rowData.values(effectiveValue,userEnteredFormat)",
    })

    const rowData = res.data.sheets?.[0]?.data?.[0]?.rowData?.[0]
    const values = rowData?.values ?? []

    const cells = values.map((val, j) => {
      return new Cell({
        value: String(extractEffectiveValue(val.effectiveValue) ?? ""),
        label: columnLabel(j),
        rowIndex: appendedRowIndex,
        cellIndex: j,
        tab: this,
        row: null,
        format: val.userEnteredFormat ?? undefined,
      })
    })

    return new Row(cells, this, appendedRowIndex)
  }

  async delRow(opts: DelRowOpts): Promise<void> {
    const { start, end } = opts
    if (start < 1) throw new Error("Row number must be >= 1")
    const client = this.getClient()
    const endIndex = end ?? start

    await client.spreadsheets.batchUpdate({
      spreadsheetId: this.getSheetId(),
      requestBody: {
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId: this.worksheetId,
                dimension: "ROWS",
                startIndex: start - 1,
                endIndex: endIndex,
              },
            },
          },
        ],
      },
    })
  }

  [Symbol.for("nodejs.util.inspect.custom")](): string {
    const rows = this._gridProperties.rowCount ?? "?"
    const cols = this._gridProperties.columnCount ?? "?"
    return `Tab(title="${this.title}", rows=${rows}, cols=${cols})`
  }

  async delCell(opts: DelCellOpts): Promise<void> {
    const { start, end, shift = "up" } = opts
    const client = this.getClient()
    const parsed = parseCellAddress(start)
    if (!parsed) throw new Error(`Invalid cell address: ${start}`)
    const startColIndex = columnIndex(parsed.label)
    const startRow = parsed.row

    let endColIndex = startColIndex
    let endRowIndex = startRow

    if (end) {
      const endParsed = parseCellAddress(end)
      if (!endParsed) throw new Error(`Invalid cell address: ${end}`)
      endColIndex = columnIndex(endParsed.label)
      endRowIndex = endParsed.row
    }

    await client.spreadsheets.batchUpdate({
      spreadsheetId: this.getSheetId(),
      requestBody: {
        requests: [
          {
            deleteRange: {
              range: {
                sheetId: this.worksheetId,
                startRowIndex: startRow - 1,
                endRowIndex: endRowIndex,
                startColumnIndex: startColIndex,
                endColumnIndex: endColIndex + 1,
              },
              shiftDimension: shift === "left" ? "COLUMNS" : "ROWS",
            },
          },
        ],
      },
    })
  }
}

/**
 * Extract the 1-based row index from a values.append updatedRange.
 * "Sheet1!A3:B3" → 3, "'My Sheet'!A3" → 3. Returns null if unparseable.
 */
function parseAppendedRowIndex(updatedRange: string): number | null {
  const afterSheet = updatedRange.split("!").pop() ?? ""
  const startCell = afterSheet.split(":")[0] ?? ""
  const m = startCell.match(/\d+/)
  return m ? parseInt(m[0], 10) : null
}

function extractEffectiveValue(
  ev: sheets_v4.Schema$ExtendedValue | undefined,
): string | number | boolean {
  if (!ev) return ""
  if (ev.stringValue != null) return ev.stringValue
  if (ev.numberValue != null) return ev.numberValue
  if (ev.boolValue != null) return ev.boolValue
  if (ev.formulaValue != null) return ev.formulaValue
  return ""
}
