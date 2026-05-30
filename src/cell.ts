import type { sheets_v4 } from "@googleapis/sheets"
import type { TabInstance, RowInstance } from "./types"
import { Style } from "./style"

type CellFormat = sheets_v4.Schema$CellFormat

export interface CellOptions {
  value: string
  label: string
  rowIndex: number
  cellIndex: number
  tab: TabInstance
  row: RowInstance | null
}

export interface CellUpdateOptions {
  inputFormat?: "raw" | "user_entered"
  renderFormat?: "formatted" | "unformatted" | "formula"
}

const VALUE_RENDER_OPTION_MAP: Record<string, string> = {
  formatted: "FORMATTED_VALUE",
  unformatted: "UNFORMATTED_VALUE",
  formula: "FORMULA",
}

const INPUT_OPTION_MAP: Record<string, string> = {
  raw: "RAW",
  user_entered: "USER_ENTERED",
}

export class Cell {
  readonly value: string
  readonly label: string
  readonly rowIndex: number
  readonly cellIndex: number
  readonly tab: TabInstance
  readonly row: RowInstance | null

  constructor(opts: CellOptions) {
    this.value = opts.value
    this.label = opts.label
    this.rowIndex = opts.rowIndex
    this.cellIndex = opts.cellIndex
    this.tab = opts.tab
    this.row = opts.row
  }

  /** @internal Set parent Row reference. Called by Row constructor. */
  _setRow(row: RowInstance): void {
    ;(this as { row: RowInstance | null }).row = row
  }

  toString(): string {
    return this.value
  }

  get header(): string {
    const headers = this.tab.getHeaders()
    return headers[this.cellIndex] ?? ""
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

  async update(newValue: string, opts: CellUpdateOptions = {}): Promise<Cell> {
    const client = this.tab.getClient()
    const range = `${this.label}${this.rowIndex}`

    await client.spreadsheets.values.update({
      spreadsheetId: this.tab.getSheetId(),
      range: `${this.tab.getTitle()}!${range}`,
      valueInputOption: INPUT_OPTION_MAP[opts.inputFormat ?? "raw"],
      requestBody: {
        values: [[newValue]],
      },
    })

    if (opts.renderFormat && opts.renderFormat !== "formatted") {
      const res = await client.spreadsheets.values.get({
        spreadsheetId: this.tab.getSheetId(),
        range: `${this.tab.getTitle()}!${range}`,
        valueRenderOption: VALUE_RENDER_OPTION_MAP[opts.renderFormat],
      })

      const rawValue = (res.data.values?.[0]?.[0] as string) ?? ""
      return new Cell({
        value: rawValue,
        label: this.label,
        rowIndex: this.rowIndex,
        cellIndex: this.cellIndex,
        tab: this.tab,
        row: this.row,
      })
    }

    return new Cell({
      value: newValue,
      label: this.label,
      rowIndex: this.rowIndex,
      cellIndex: this.cellIndex,
      tab: this.tab,
      row: this.row,
    })
  }

  async clear(): Promise<void> {
    const client = this.tab.getClient()
    const range = `${this.label}${this.rowIndex}`

    await client.spreadsheets.values.clear({
      spreadsheetId: this.tab.getSheetId(),
      range: `${this.tab.getTitle()}!${range}`,
    })
  }

  async style(obj: Style | CellFormat): Promise<void> {
    const client = this.tab.getClient()
    const cellFormat = obj instanceof Style ? obj.toCellFormat() : obj

    await client.spreadsheets.batchUpdate({
      spreadsheetId: this.tab.getSheetId(),
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId: this.tab.getWorksheetId(),
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
    const client = this.tab.getClient()
    const isLeft = shift === "left"

    await client.spreadsheets.batchUpdate({
      spreadsheetId: this.tab.getSheetId(),
      requestBody: {
        requests: [
          {
            deleteRange: {
              range: {
                sheetId: this.tab.getWorksheetId(),
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
