/** Convert 0-based column index to label: 0 → "A", 25 → "Z", 26 → "AA" */
export function columnLabel(index: number): string {
  let label = ""
  let n = index
  while (n >= 0) {
    label = String.fromCharCode(65 + (n % 26)) + label
    n = Math.floor(n / 26) - 1
  }
  return label
}

/** Convert column label to 0-based index: "A" → 0, "Z" → 25, "AA" → 26 */
export function columnIndex(label: string): number {
  return (
    label
      .split("")
      .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1
  )
}

export interface CellAddress {
  label: string
  row: number
}

/** Parse A1 cell address: "B2" → { label: "B", row: 2 }. Returns null on invalid input. */
export function parseCellAddress(address: string): CellAddress | null {
  const m = address.match(/^([A-Z]+)(\d+)$/)
  if (!m) return null
  return { label: m[1] ?? "", row: parseInt(m[2] ?? "0", 10) }
}

import type { sheets_v4 } from "@googleapis/sheets"
import { ZodError } from "zod"
import { Cell } from "./cell"
import { ValidationError, type ValueInputOption } from "./types"
import type { z } from "zod"

type CellFormat = sheets_v4.Schema$CellFormat

/** Extract string value from a Cell object or raw value. Shared by append/update. */
export function extractCellValue(v: unknown): string {
  return v instanceof Cell ? v.value : String(v ?? "")
}

/** Extract CellFormat from a Cell object's Format. Undefined if no format. */
export function extractCellFormat(v: unknown): CellFormat | undefined {
  return v instanceof Cell && v.format ? v.format.toCellFormat() : undefined
}

/** Build a Google API cell data object from value + optional format. */
export function buildCellData(
  v: unknown,
  valueInputOption: ValueInputOption = "RAW",
): {
  userEnteredValue: { stringValue: string } | { formulaValue: string }
  userEnteredFormat?: CellFormat
} {
  const value = extractCellValue(v)
  const format = extractCellFormat(v)
  const userEnteredValue =
    valueInputOption === "USER_ENTERED" && value.startsWith("=")
      ? { formulaValue: value }
      : { stringValue: value }
  const cell: {
    userEnteredValue: { stringValue: string } | { formulaValue: string }
    userEnteredFormat?: CellFormat
  } = { userEnteredValue }
  if (format) cell.userEnteredFormat = format
  return cell
}

/**
 * Validate array values against a Zod schema by column position.
 * Throws ValidationError on mismatch. Shared by tab.append / row.update.
 */
export function validateArrayAgainstSchema(
  values: unknown[],
  headers: string[],
  schema: z.ZodObject,
): void {
  for (let i = 0; i < values.length; i++) {
    const colIdx = i
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

/** Convert hex color string to Google Sheets Color object. */
export function hexToColor(
  hex: string,
): { red: number; green: number; blue: number } | null {
  const m = hex
    .replace("#", "")
    .match(/^([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/)
  if (!m) return null
  const r = m[1]
  const g = m[2]
  const b = m[3]
  if (r === undefined || g === undefined || b === undefined) return null
  return {
    red: parseInt(r, 16) / 255,
    green: parseInt(g, 16) / 255,
    blue: parseInt(b, 16) / 255,
  }
}
