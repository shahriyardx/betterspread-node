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
  if (label.length === 0) return -1
  const upper = label.toUpperCase()
  for (const char of upper) {
    const code = char.charCodeAt(0)
    if (code < 65 || code > 90) return -1
  }
  return (
    upper
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
    validateFieldWithCoercion(fieldSchema, val, header)
  }
}

/**
 * Candidate coercions for a raw string value: number, boolean, date.
 * Used to reconcile typed object validation with string array validation.
 */
function tryCoercions(str: string): unknown[] {
  const candidates: unknown[] = []
  if (str.trim() !== "") {
    const num = Number(str)
    if (!Number.isNaN(num)) candidates.push(num)
  }
  const lower = str.toLowerCase()
  if (lower === "true") candidates.push(true)
  else if (lower === "false") candidates.push(false)
  const date = new Date(str)
  if (!Number.isNaN(date.getTime())) candidates.push(date)
  return candidates
}

/**
 * Validate a single value against a field schema, attempting string coercions
 * (number/boolean/date) before failing. Throws ValidationError if all fail.
 */
export function validateFieldWithCoercion(
  fieldSchema: z.ZodType,
  val: unknown,
  header: string,
): void {
  const first = fieldSchema.safeParse(val)
  if (first.success) return
  if (typeof val === "string") {
    for (const candidate of tryCoercions(val)) {
      if (fieldSchema.safeParse(candidate).success) return
    }
  }
  throw new ValidationError(
    `Schema validation failed for column "${header}": ${first.error.message}`,
    first.error,
  )
}

/**
 * Convert hex color string to Google Sheets Color object.
 * Supports 6-digit (#FF00FF) and 3-digit shorthand (#F0F), with or without "#".
 * @throws Error if the input is not a valid hex color.
 */
export function hexToColor(hex: string): {
  red: number
  green: number
  blue: number
} {
  let normalized = hex.trim().replace(/^#/, "")
  if (/^[0-9a-fA-F]{3}$/.test(normalized)) {
    normalized = normalized
      .split("")
      .map((c) => c + c)
      .join("")
  }
  const m = normalized.match(
    /^([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/,
  )
  if (!m) {
    throw new Error(`Invalid hex color: ${hex}`)
  }
  const r = m[1] as string
  const g = m[2] as string
  const b = m[3] as string
  return {
    red: parseInt(r, 16) / 255,
    green: parseInt(g, 16) / 255,
    blue: parseInt(b, 16) / 255,
  }
}
