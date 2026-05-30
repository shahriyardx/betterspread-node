import type { sheets_v4 } from "@googleapis/sheets"

type CellFormat = sheets_v4.Schema$CellFormat

export interface StyleOptions {
  bgColor?: string
  textColor?: string
  horizontalAlign?: "left" | "center" | "right"
  verticalAlign?: "top" | "middle" | "bottom"
  bold?: boolean
  italic?: boolean
  strikethrough?: boolean
  raw?: CellFormat
}

const HORIZONTAL_MAP: Record<string, string> = {
  left: "LEFT",
  center: "CENTER",
  right: "RIGHT",
}

const VERTICAL_MAP: Record<string, string> = {
  top: "TOP",
  middle: "MIDDLE",
  bottom: "BOTTOM",
}



function hexToColor(
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

export class Style {
  readonly bgColor?: string
  readonly textColor?: string
  readonly horizontalAlign?: "left" | "center" | "right"
  readonly verticalAlign?: "top" | "middle" | "bottom"
  readonly bold?: boolean
  readonly italic?: boolean
  readonly strikethrough?: boolean
  readonly raw?: CellFormat

  constructor(opts: StyleOptions = {}) {
    this.bgColor = opts.bgColor
    this.textColor = opts.textColor
    this.horizontalAlign = opts.horizontalAlign
    this.verticalAlign = opts.verticalAlign
    this.bold = opts.bold
    this.italic = opts.italic
    this.strikethrough = opts.strikethrough
    this.raw = opts.raw
  }

  toCellFormat(): CellFormat {
    if (this.raw) return this.raw

    const format: CellFormat = {}

    if (this.bgColor) {
      const c = hexToColor(this.bgColor)
      if (c) format.backgroundColor = c
    }

    if (this.textColor) {
      const c = hexToColor(this.textColor)
      if (c) {
        format.textFormat = { ...format.textFormat, foregroundColor: c }
      }
    }

    if (
      this.bold !== undefined ||
      this.italic !== undefined ||
      this.strikethrough !== undefined
    ) {
      format.textFormat = {
        ...format.textFormat,
        bold: this.bold,
        italic: this.italic,
        strikethrough: this.strikethrough,
      }
    }

    if (this.horizontalAlign) {
      format.horizontalAlignment = HORIZONTAL_MAP[this.horizontalAlign]
    }

    if (this.verticalAlign) {
      format.verticalAlignment = VERTICAL_MAP[this.verticalAlign]
    }

    return format
  }
}
