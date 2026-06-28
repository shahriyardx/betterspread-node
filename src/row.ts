import { ZodError } from "zod"
import { Cell } from "./cell"
import type { TabInstance, RowInstance, ValueInputOption } from "./types"
import { ValidationError } from "./types"
import type { Format } from "./format"
import {
  columnLabel,
  columnIndex,
  buildCellData,
  extractCellValue,
  extractCellFormat,
  validateArrayAgainstSchema,
} from "./utils"

export interface RowUpdateOpts {
  values: unknown[] | Record<string, unknown>
  inputFormat?: ValueInputOption
}

export class Row extends Array<Cell> implements RowInstance {
  private _tab: TabInstance
  private _rowIndex: number

  static override get [Symbol.species](): ArrayConstructor {
    return Array
  }

  constructor(cells: Cell[], tab: TabInstance, rowIndex: number) {
    super(...cells)
    this._tab = tab
    this._rowIndex = rowIndex

    const proxy = new Proxy(this, {
      get(target, prop) {
        if (typeof prop === "string" && /^\d+$/.test(prop)) {
          const idx = parseInt(prop, 10)
          if (idx >= 0 && idx < target.length) {
            return target[idx]
          }
          return new Cell({
            value: "",
            label: columnLabel(idx),
            rowIndex: target._rowIndex,
            cellIndex: idx,
            tab: target._tab,
            row: proxy,
          })
        }
        return Reflect.get(target, prop)
      },
    })

    for (const cell of cells) {
      cell._setRow(proxy)
    }

    // biome-ignore lint/correctness/noConstructorReturn: Proxy wraps Row for virtual cell access
    return proxy
  }

  [Symbol.for("nodejs.util.inspect.custom")](): string {
    return `Row(index=${this._rowIndex}, cells=${this.length})`
  }

  getTab(): TabInstance {
    return this._tab
  }

  getRowIndex(): number {
    return this._rowIndex
  }

  get(col: string): Cell {
    const idx = columnIndex(col.toUpperCase())
    const existing = this[idx]
    if (existing) return existing

    return new Cell({
      value: "",
      label: col.toUpperCase(),
      rowIndex: this._rowIndex,
      cellIndex: idx,
      tab: this._tab,
      row: this,
    })
  }

  private objectToRowUpdate(
    data: Record<string, unknown>,
  ): { colIndex: number; value: unknown }[] {
    const headers = this._tab.getHeaders()
    if (headers.length === 0) {
      throw new Error("Tab headers not cached. Call values() first.")
    }

    const schema = this._tab.getSchema()
    if (schema) {
      try {
        schema.parse(data)
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

    const pairs: { colIndex: number; value: unknown }[] = []
    for (const [key, val] of Object.entries(data)) {
      const idx = headers.indexOf(key)
      if (idx === -1) continue
      pairs.push({ colIndex: idx, value: val })
    }
    pairs.sort((a, b) => a.colIndex - b.colIndex)
    return pairs
  }

  async update(opts: RowUpdateOpts): Promise<void> {
    const client = this._tab.getClient()
    const { values, inputFormat } = opts

    // Validate against schema if set
    const schema = this._tab.getSchema()
    const headers = this._tab.getHeaders()
    if (schema && headers.length > 0 && Array.isArray(values)) {
      validateArrayAgainstSchema(values, headers, schema)
    }

    // Convert object to column-indexed pairs
    const colValPairs = Array.isArray(values)
      ? values.map((v, i) => ({ colIndex: i, value: v }))
      : this.objectToRowUpdate(values)

    if (colValPairs.length === 0) return

    // Build contiguous row data for updateCells
    const firstPair = colValPairs[0]
    const lastPair = colValPairs[colValPairs.length - 1]
    if (!firstPair || !lastPair) return

    const minCol = firstPair.colIndex
    const maxCol = lastPair.colIndex
    const cellData: Record<string, unknown>[] = []
    let pairIdx = 0
    for (let c = minCol; c <= maxCol; c++) {
      const currentPair = colValPairs[pairIdx]
      if (currentPair && currentPair.colIndex === c) {
        cellData.push(
          buildCellData(currentPair.value, inputFormat) as Record<
            string,
            unknown
          >,
        )
        pairIdx++
      } else {
        cellData.push({ userEnteredValue: { stringValue: "" } })
      }
    }

    await client.spreadsheets.batchUpdate({
      spreadsheetId: this._tab.getSheetId(),
      requestBody: {
        requests: [
          {
            updateCells: {
              range: {
                sheetId: this._tab.getWorksheetId(),
                startRowIndex: this._rowIndex - 1,
                endRowIndex: this._rowIndex,
                startColumnIndex: minCol,
                endColumnIndex: maxCol + 1,
              },
              rows: [{ values: cellData }],
              fields: "userEnteredValue,userEnteredFormat",
            },
          },
        ],
      },
    })

    // Update local cells
    for (const p of colValPairs) {
      const val = p.value
      this[p.colIndex] = new Cell({
        value: extractCellValue(val),
        label: columnLabel(p.colIndex),
        rowIndex: this._rowIndex,
        cellIndex: p.colIndex,
        tab: this._tab,
        row: this,
        format: extractCellFormat(val),
      })
    }
  }

  async clear(): Promise<void> {
    if (this.length === 0) return
    const client = this._tab.getClient()
    const endCol = columnLabel(this.length - 1)
    const range = `A${this._rowIndex}:${endCol}${this._rowIndex}`

    await client.spreadsheets.values.clear({
      spreadsheetId: this._tab.getSheetId(),
      range: `${this._tab.getTitle()}!${range}`,
    })
  }

  async style(style: Format): Promise<void> {
    const client = this._tab.getClient()
    const cellFormat = style.toCellFormat()

    await client.spreadsheets.batchUpdate({
      spreadsheetId: this._tab.getSheetId(),
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId: this._tab.getWorksheetId(),
                startRowIndex: this._rowIndex - 1,
                endRowIndex: this._rowIndex,
              },
              cell: {
                userEnteredFormat: cellFormat,
              },
              fields: "userEnteredFormat",
            },
          },
        ],
      },
    })
  }

  async appendCell(value: unknown | unknown[]): Promise<void> {
    const client = this._tab.getClient()
    const vals = Array.isArray(value) ? value : [value]
    const startCol = columnLabel(this.length)
    const range = `${startCol}${this._rowIndex}`

    await client.spreadsheets.values.update({
      spreadsheetId: this._tab.getSheetId(),
      range: `${this._tab.getTitle()}!${range}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [vals],
      },
    })

    // Update local cells
    const startIdx = this.length
    const newCells = vals.map((v, i) => {
      return new Cell({
        value: extractCellValue(v),
        label: columnLabel(startIdx + i),
        rowIndex: this._rowIndex,
        cellIndex: startIdx + i,
        tab: this._tab,
        row: this,
        format: extractCellFormat(v),
      })
    })
    this.push(...newCells)
  }

  async refetch(): Promise<void> {
    const client = this._tab.getClient()
    const endCol = columnLabel(this.length > 0 ? this.length - 1 : 0)
    const range = `A${this._rowIndex}:${endCol}${this._rowIndex}`

    const res = await client.spreadsheets.values.get({
      spreadsheetId: this._tab.getSheetId(),
      range: `${this._tab.getTitle()}!${range}`,
    })

    const rawValues = (res.data.values?.[0] as unknown[]) ?? []
    const newCells = rawValues.map((val, i) => {
      return new Cell({
        value: String(val ?? ""),
        label: columnLabel(i),
        rowIndex: this._rowIndex,
        cellIndex: i,
        tab: this._tab,
        row: this,
      })
    })

    // Replace contents
    this.length = 0
    this.push(...newCells)
  }

  async delete(): Promise<void> {
    const client = this._tab.getClient()

    await client.spreadsheets.batchUpdate({
      spreadsheetId: this._tab.getSheetId(),
      requestBody: {
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId: this._tab.getWorksheetId(),
                dimension: "ROWS",
                startIndex: this._rowIndex - 1,
                endIndex: this._rowIndex,
              },
            },
          },
        ],
      },
    })
  }
}
