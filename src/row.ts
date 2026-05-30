import type { sheets_v4 } from "@googleapis/sheets"
import { ZodError } from "zod"
import { Cell } from "./cell"
import type { TabInstance, RowInstance } from "./types"
import { ValidationError } from "./types"
import { Style } from "./style"
import { columnLabel, columnIndex } from "./utils"

type CellFormat = sheets_v4.Schema$CellFormat

export class Row extends Array<Cell> implements RowInstance {
  private _tab: TabInstance
  private _rowIndex: number

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

  async update(values: unknown[] | Record<string, unknown>, start: string = "A"): Promise<void> {
    const client = this._tab.getClient()

    // Validate against schema if set
    const schema = this._tab.getSchema()
    const headers = this._tab.getHeaders()
    if (schema && headers.length > 0 && Array.isArray(values)) {
      for (let i = 0; i < values.length; i++) {
        const colIdx = i + (start.charCodeAt(0) - 65)
        const header = headers[colIdx]
        if (!header) continue
        const fieldSchema = schema.shape[header]
        if (!fieldSchema) continue
        const item = values[i]
        const val = item instanceof Cell ? item.value : item
        try {
          fieldSchema.parse(val)
        } catch (err) {
          if (err instanceof ZodError) {
            throw new ValidationError(
              `Schema validation failed for column "${header}": ${err.message}`,
              err,
            )
          }
          throw err
        }
      }
    }

    // Convert object to column-indexed pairs
    const colValPairs = Array.isArray(values)
      ? values.map((v, i) => ({ colIndex: i + (start.charCodeAt(0) - 65), value: v }))
      : this.objectToRowUpdate(values)

    if (colValPairs.length === 0) return

    // Build contiguous array from min to max col for API call
    const firstPair = colValPairs[0]
    const lastPair = colValPairs[colValPairs.length - 1]
    if (!firstPair || !lastPair) return

    const minCol = firstPair.colIndex
    const maxCol = lastPair.colIndex
    const rawValues: unknown[] = []
    let pairIdx = 0
    for (let c = minCol; c <= maxCol; c++) {
      const currentPair = colValPairs[pairIdx]
      if (currentPair && currentPair.colIndex === c) {
        const v = currentPair.value
        rawValues.push(v instanceof Cell ? v.value : v)
        pairIdx++
      } else {
        rawValues.push("")
      }
    }

    const startLabel = columnLabel(minCol)
    const range = `${startLabel}${this._rowIndex}`

    await client.spreadsheets.values.update({
      spreadsheetId: this._tab.getSheetId(),
      range: `${this._tab.getTitle()}!${range}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [rawValues],
      },
    })

    // Apply formatting from Cell objects
    const formatRequests = colValPairs
      .filter((p) => {
        const cell = p.value
        return cell instanceof Cell && cell.format !== null
      })
      .map((p) => {
        const cell = p.value as Cell
        return {
          repeatCell: {
            range: {
              sheetId: this._tab.getWorksheetId(),
              startRowIndex: this._rowIndex - 1,
              endRowIndex: this._rowIndex,
              startColumnIndex: p.colIndex,
              endColumnIndex: p.colIndex + 1,
            },
            cell: { userEnteredFormat: cell.format ?? undefined },
            fields: "userEnteredFormat",
          },
        }
      })

    if (formatRequests.length > 0) {
      await client.spreadsheets.batchUpdate({
        spreadsheetId: this._tab.getSheetId(),
        requestBody: { requests: formatRequests },
      })
    }

    // Update local cells
    for (const p of colValPairs) {
      if (p.colIndex < this.length) {
        const val = p.value
        this[p.colIndex] = new Cell({
          value: val instanceof Cell ? val.value : String(val ?? ""),
          label: columnLabel(p.colIndex),
          rowIndex: this._rowIndex,
          cellIndex: p.colIndex,
          tab: this._tab,
          row: this,
          format: val instanceof Cell ? (val.format ?? undefined) : undefined,
        })
      }
    }
  }

  async clear(): Promise<void> {
    const client = this._tab.getClient()
    const endCol = columnLabel(this.length - 1)
    const range = `A${this._rowIndex}:${endCol}${this._rowIndex}`

    await client.spreadsheets.values.clear({
      spreadsheetId: this._tab.getSheetId(),
      range: `${this._tab.getTitle()}!${range}`,
    })
  }

  async style(obj: Style | CellFormat): Promise<void> {
    const client = this._tab.getClient()
    const cellFormat = obj instanceof Style ? obj.toCellFormat() : obj

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
