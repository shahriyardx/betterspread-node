import type { sheets_v4 } from "@googleapis/sheets"
import { Cell } from "./cell"
import type { TabInstance, RowInstance } from "./types"
import { Style } from "./style"
import { columnLabel } from "./utils"

type CellFormat = sheets_v4.Schema$CellFormat

export class Row extends Array<Cell> implements RowInstance {
  private _tab: TabInstance
  private _rowIndex: number

  constructor(cells: Cell[], tab: TabInstance, rowIndex: number) {
    super(...cells)
    this._tab = tab
    this._rowIndex = rowIndex
    for (const cell of cells) {
      cell._setRow(this)
    }
  }

  getTab(): TabInstance {
    return this._tab
  }

  getRowIndex(): number {
    return this._rowIndex
  }

  async update(values: unknown[], start: string = "A"): Promise<void> {
    const client = this._tab.getClient()
    const startCol = start.toUpperCase().charCodeAt(0) - 65
    const range = `${start}${this._rowIndex}`

    await client.spreadsheets.values.update({
      spreadsheetId: this._tab.getSheetId(),
      range: `${this._tab.getTitle()}!${range}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [values],
      },
    })

    // Update local cells
    for (let i = 0; i < values.length; i++) {
      const cellIndex = startCol + i
      if (cellIndex < this.length) {
        this[cellIndex] = new Cell({
          value: String(values[i] ?? ""),
          label: columnLabel(cellIndex),
          rowIndex: this._rowIndex,
          cellIndex,
          tab: this._tab,
          row: this,
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
