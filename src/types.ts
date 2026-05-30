import type { sheets_v4 } from "@googleapis/sheets"

export class ValueError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "ValueError"
  }
}

export class SheetNotFoundError extends Error {
  constructor(name: string) {
    super(`Sheet "${name}" not found`)
    this.name = "SheetNotFoundError"
  }
}

export class TabNotFoundError extends Error {
  constructor(name: string) {
    super(`Tab "${name}" not found`)
    this.name = "TabNotFoundError"
  }
}

export interface CellProps {
  value: string
  label: string
  rowIndex: number
  cellIndex: number
  tab: TabInstance
  row: RowInstance | null
}

export interface TabInstance {
  getClient(): sheets_v4.Sheets
  getSheetId(): string
  getTitle(): string
  getWorksheetId(): number
  getHeaders(): string[]
}

export interface RowInstance {
  getTab(): TabInstance
  getRowIndex(): number
}
