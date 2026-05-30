import type { sheets_v4 } from "@googleapis/sheets"
import { ZodError } from "zod"
import type { TabInstance, RowInstance } from "./types"
import { ValidationError } from "./types"
import { Format } from "./format"

type CellFormat = sheets_v4.Schema$CellFormat

export interface CellOptions {
  value: string
  label?: string
  rowIndex?: number
  cellIndex?: number
  tab?: TabInstance
  row?: RowInstance | null
  format?: CellFormat | Format
}

export class Cell {
  readonly value: string
  readonly label: string
  readonly rowIndex: number
  readonly cellIndex: number
  readonly tab?: TabInstance
  readonly row: RowInstance | null
  format: Format | null

  constructor(opts: CellOptions) {
    this.value = opts.value
    this.label = opts.label ?? "A"
    this.rowIndex = opts.rowIndex ?? 1
    this.cellIndex = opts.cellIndex ?? 0
    this.tab = opts.tab
    this.row = opts.row ?? null
    this.format =
      opts.format instanceof Format
        ? opts.format
        : opts.format
          ? Format.fromCellFormat(opts.format)
          : null
  }

  private requireTab(): TabInstance {
    if (!this.tab) throw new Error("Cell has no tab reference")
    return this.tab
  }

  /** @internal Set parent Row reference. Called by Row constructor. */
  _setRow(row: RowInstance): void {
    ;(this as { row: RowInstance | null }).row = row
  }

  toString(): string {
    return this.value
  }

  get header(): string {
    const headers = this.tab?.getHeaders()
    return headers?.[this.cellIndex] ?? ""
  }

  toJSON(): string {
    return this.value
  }

  [Symbol.toPrimitive](hint: string): string | number {
    if (hint === "number") return Number(this.value) || 0
    return this.value
  }

  [Symbol.for("nodejs.util.inspect.custom")](): string {
    return `Cell(row=${this.rowIndex}, header="${this.header}", value="${this.value}")`
  }

  async update(newValue: string | Cell, format?: Format): Promise<this> {
    const tab = this.requireTab()
    const client = tab.getClient()

    // Resolve value and optional format
    const isCell = newValue instanceof Cell
    const strValue = isCell ? newValue.value : newValue
    const cellFormat = isCell
      ? (newValue.format?.toCellFormat() ?? undefined)
      : (format?.toCellFormat() ?? undefined)

    // Validate against schema if applicable
    const schema = tab.getSchema()
    if (schema) {
      const header = this.header
      if (header && header in schema.shape) {
        const fieldSchema = schema.shape[header]
        try {
          fieldSchema.parse(strValue)
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

    // Build cell data and push via batchUpdate
    const cellData: Record<string, unknown> = {
      userEnteredValue: { stringValue: strValue },
    }
    const fields = ["userEnteredValue"]
    if (cellFormat) {
      cellData.userEnteredFormat = cellFormat
      fields.push("userEnteredFormat")
    }

    await client.spreadsheets.batchUpdate({
      spreadsheetId: tab.getSheetId(),
      requestBody: {
        requests: [
          {
            updateCells: {
              range: {
                sheetId: tab.getWorksheetId(),
                startRowIndex: this.rowIndex - 1,
                endRowIndex: this.rowIndex,
                startColumnIndex: this.cellIndex,
                endColumnIndex: this.cellIndex + 1,
              },
              rows: [{ values: [cellData] }],
              fields: fields.join(","),
            },
          },
        ],
      },
    })

    // Update local state
    ;(this as { value: string }).value = strValue
    if (cellFormat) this.format = Format.fromCellFormat(cellFormat)

    return this
  }

  async clear(): Promise<void> {
    const tab = this.requireTab()
    const client = tab.getClient()
    const range = `${this.label}${this.rowIndex}`

    await client.spreadsheets.values.clear({
      spreadsheetId: tab.getSheetId(),
      range: `${tab.getTitle()}!${range}`,
    })
  }

  async style(style: Format): Promise<void> {
    const tab = this.requireTab()
    const client = tab.getClient()
    const cellFormat = style.toCellFormat()

    await client.spreadsheets.batchUpdate({
      spreadsheetId: tab.getSheetId(),
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId: tab.getWorksheetId(),
                startRowIndex: this.rowIndex - 1,
                endRowIndex: this.rowIndex,
                startColumnIndex: this.cellIndex,
                endColumnIndex: this.cellIndex + 1,
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

  async delete(shift: "left" | "up" = "left"): Promise<void> {
    const tab = this.requireTab()
    const client = tab.getClient()
    const isLeft = shift === "left"

    await client.spreadsheets.batchUpdate({
      spreadsheetId: tab.getSheetId(),
      requestBody: {
        requests: [
          {
            deleteRange: {
              range: {
                sheetId: tab.getWorksheetId(),
                startRowIndex: this.rowIndex - 1,
                endRowIndex: this.rowIndex,
                startColumnIndex: this.cellIndex,
                endColumnIndex: this.cellIndex + 1,
              },
              shiftDimension: isLeft ? "COLUMNS" : "ROWS",
            },
          },
        ],
      },
    })
  }
}
