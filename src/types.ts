import type { sheets_v4 } from "@googleapis/sheets"
import type { z } from "zod"

export class ValueError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "ValueError"
  }
}

export class ValidationError extends Error {
  constructor(message: string, cause?: unknown) {
    super(message)
    this.name = "ValidationError"
    this.cause = cause
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

export interface TabInstance {
  getClient(): sheets_v4.Sheets
  getSheetId(): string
  getTitle(): string
  getWorksheetId(): number
  getHeaders(): string[]
  getSchema(): z.ZodObject | null
}

export interface RowInstance {
  getTab(): TabInstance
  getRowIndex(): number
}
