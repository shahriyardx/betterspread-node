import type { sheets_v4 } from "@googleapis/sheets"
import { Cell } from "./cell"
import { Row } from "./row"
import type { Sheet } from "./sheet"
import { columnLabel, columnIndex, parseCellAddress } from "./utils"

export class Tab {
  private sheet: Sheet
  private title: string
  private worksheetId: number
  private _gridProperties: sheets_v4.Schema$GridProperties
  private _headers: string[] = []

  constructor(
    sheet: Sheet,
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
    this._headers = headers
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

  async values(range?: string): Promise<Row[]> {
    const client = this.getClient()
    const actualRange = range ?? `${this.title}!A1:ZZZ`

    const res = await client.spreadsheets.values.get({
      spreadsheetId: this.getSheetId(),
      range: actualRange,
    })

    const rows = (res.data.values as unknown[][]) ?? []

    // Cache headers from first row
    const firstRow = rows[0]
    if (firstRow) {
      this._headers = firstRow.map((h) => String(h ?? ""))
    }

    return rows.map((rowValues, i) => {
      const cells = rowValues.map((val, j) => {
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

  async getRow(serialNo: number): Promise<Row> {
    const client = this.getClient()
    const range = `${this.title}!A${serialNo}:ZZZ${serialNo}`

    const res = await client.spreadsheets.values.get({
      spreadsheetId: this.getSheetId(),
      range,
    })

    const values = (res.data.values?.[0] as unknown[]) ?? []

    if (values.length === 0) {
      return new Row([], this, serialNo)
    }

    const cells = values.map((val, j) => {
      return new Cell({
        value: String(val ?? ""),
        label: columnLabel(j),
        rowIndex: serialNo,
        cellIndex: j,
        tab: this,
        row: null,
      })
    })

    return new Row(cells, this, serialNo)
  }

  async getCell(
    cellName: string,
    renderOption: "formatted" | "unformatted" | "formula" = "formatted",
  ): Promise<Cell> {
    const client = this.getClient()

    const valueRenderOptionMap: Record<string, string> = {
      formatted: "FORMATTED_VALUE",
      unformatted: "UNFORMATTED_VALUE",
      formula: "FORMULA",
    }

    const res = await client.spreadsheets.values.get({
      spreadsheetId: this.getSheetId(),
      range: `${this.title}!${cellName}`,
      valueRenderOption: valueRenderOptionMap[renderOption],
    })

    const value = (res.data.values?.[0]?.[0] as string) ?? ""
    const addr = parseCellAddress(cellName)
    const label = addr?.label ?? "A"
    const cellIndex = columnIndex(label)
    const rowIndex = addr?.row ?? parseInt(cellName.match(/\d+/)?.[0] ?? "1", 10)

    return new Cell({
      value,
      label,
      rowIndex,
      cellIndex,
      tab: this,
      row: null,
    })
  }

  async append(data: unknown[], getRow: boolean = false): Promise<Row | null> {
    const client = this.getClient()

    await client.spreadsheets.values.append({
      spreadsheetId: this.getSheetId(),
      range: `${this.title}!A:A`,
      valueInputOption: "RAW",
      requestBody: {
        values: [data],
      },
    })

    if (!getRow) return null

    // Refetch to find the last row
    const res = await client.spreadsheets.values.get({
      spreadsheetId: this.getSheetId(),
      range: `${this.title}!A:ZZZ`,
    })

    const rows = (res.data.values as unknown[][]) ?? []
    const lastRowValues = rows[rows.length - 1] ?? []

    const cells = lastRowValues.map((val, j) => {
      return new Cell({
        value: String(val ?? ""),
        label: columnLabel(j),
        rowIndex: rows.length,
        cellIndex: j,
        tab: this,
        row: null,
      })
    })

    return new Row(cells, this, rows.length)
  }

  async delRow(start: number, end?: number): Promise<void> {
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

  async delCell(
    start: string,
    end?: string,
    shift: "up" | "left" = "up",
  ): Promise<void> {
    const client = this.getClient()
    const parsed = parseCellAddress(start)
    if (!parsed) throw new Error(`Invalid cell address: ${start}`)
    const startColIndex = columnIndex(parsed.label)
    const startRow = parsed.row

    let endColIndex = startColIndex
    let endRowIndex = startRow

    if (end) {
      const endParsed = parseCellAddress(end)
      if (endParsed) {
        endColIndex = columnIndex(endParsed.label)
        endRowIndex = endParsed.row
      }
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
