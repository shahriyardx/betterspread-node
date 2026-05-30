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
  return label
    .split("")
    .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1
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
