import type { sheets_v4 } from "@googleapis/sheets"
import { hexToColor } from "./utils"

type CellFormat = sheets_v4.Schema$CellFormat
type Color = sheets_v4.Schema$Color
type ColorStyle = sheets_v4.Schema$ColorStyle
type Borders = sheets_v4.Schema$Borders
type Border = sheets_v4.Schema$Border
type Padding = sheets_v4.Schema$Padding
type NumberFormat = sheets_v4.Schema$NumberFormat
type TextFormat = sheets_v4.Schema$TextFormat
type TextRotation = sheets_v4.Schema$TextRotation
type Link = sheets_v4.Schema$Link

export interface BorderOptions {
  color?: Color
  colorStyle?: ColorStyle
  style?:
    | "solid"
    | "dashed"
    | "dotted"
    | "double"
    | "none"
    | "medium"
    | "mediumDashed"
    | "mediumDotted"
    | "mediumSolid"
    | "thick"
  width?: number
}

export interface StyleOptions {
  /** Background color as hex string (e.g. "#ff0000"). Maps to backgroundColor. */
  bgColor?: string
  /** Text color as hex string (e.g. "#000000"). Maps to textFormat.foregroundColor. */
  textColor?: string
  /** Shorthand for textFormat.bold. */
  bold?: boolean
  /** Shorthand for textFormat.italic. */
  italic?: boolean
  /** Shorthand for textFormat.strikethrough. */
  strikethrough?: boolean
  /** Shorthand for textFormat.underline. */
  underline?: boolean
  /** Shorthand for textFormat.fontFamily. */
  fontFamily?: string
  /** Shorthand for textFormat.fontSize. */
  fontSize?: number
  /** Shorthand for textFormat.link. */
  link?: Link
  /** Horizontal alignment. Mapped to CellFormat values. */
  horizontalAlign?: "left" | "center" | "right"
  /** Vertical alignment. Mapped to CellFormat values. */
  verticalAlign?: "top" | "middle" | "bottom"

  // Full CellFormat properties

  /** Background color as Color object. Overrides bgColor if both set. */
  backgroundColor?: Color
  /** Background color style. Takes precedence over backgroundColor. */
  backgroundColorStyle?: ColorStyle
  /** Cell borders. */
  borders?: {
    top?: BorderOptions
    bottom?: BorderOptions
    left?: BorderOptions
    right?: BorderOptions
  }
  /** Horizontal alignment (raw CellFormat value). Overrides horizontalAlign if both set. */
  horizontalAlignment?: string
  /** Vertical alignment (raw CellFormat value). Overrides verticalAlign if both set. */
  verticalAlignment?: string
  /** How hyperlinks display. */
  hyperlinkDisplayType?: "LINKED" | "PLAIN_TEXT"
  /** Number format. */
  numberFormat?: NumberFormat
  /** Cell padding. */
  padding?: Padding
  /** Text direction. */
  textDirection?: "LEFT_TO_RIGHT" | "RIGHT_TO_LEFT"
  /** Full text format object. Merged with shorthands. */
  textFormat?: TextFormat
  /** Text rotation. */
  textRotation?: TextRotation
  /** Wrap strategy. */
  wrapStrategy?: "OVERFLOW_CELL" | "LEGACY_WRAP" | "CLIP" | "WRAP"

  /** Raw CellFormat — bypasses all other options when set. */
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

const BORDER_STYLE_MAP: Record<string, string> = {
  solid: "SOLID",
  dashed: "DASHED",
  dotted: "DOTTED",
  double: "DOUBLE",
  none: "NONE",
  medium: "MEDIUM",
  mediumDashed: "MEDIUM_DASHED",
  mediumDotted: "MEDIUM_DOTTED",
  mediumSolid: "MEDIUM_SOLID",
  thick: "THICK",
}

function toCellFormatBorder(border?: BorderOptions): Border | undefined {
  if (!border) return undefined
  const b: Border = {}
  if (border.color) b.color = border.color
  if (border.colorStyle) b.colorStyle = border.colorStyle
  if (border.style) b.style = BORDER_STYLE_MAP[border.style]
  if (border.width !== undefined) b.width = border.width
  return Object.keys(b).length > 0 ? b : undefined
}

export class Style {
  readonly bgColor?: string
  readonly textColor?: string
  readonly bold?: boolean
  readonly italic?: boolean
  readonly strikethrough?: boolean
  readonly underline?: boolean
  readonly fontFamily?: string
  readonly fontSize?: number
  readonly link?: Link
  readonly horizontalAlign?: "left" | "center" | "right"
  readonly verticalAlign?: "top" | "middle" | "bottom"

  readonly backgroundColor?: Color
  readonly backgroundColorStyle?: ColorStyle
  readonly borders?: {
    top?: BorderOptions
    bottom?: BorderOptions
    left?: BorderOptions
    right?: BorderOptions
  }
  readonly horizontalAlignment?: string
  readonly verticalAlignment?: string
  readonly hyperlinkDisplayType?: "LINKED" | "PLAIN_TEXT"
  readonly numberFormat?: NumberFormat
  readonly padding?: Padding
  readonly textDirection?: "LEFT_TO_RIGHT" | "RIGHT_TO_LEFT"
  readonly textFormat?: TextFormat
  readonly textRotation?: TextRotation
  readonly wrapStrategy?: "OVERFLOW_CELL" | "LEGACY_WRAP" | "CLIP" | "WRAP"

  readonly raw?: CellFormat

  constructor(opts: StyleOptions = {}) {
    this.bgColor = opts.bgColor
    this.textColor = opts.textColor
    this.bold = opts.bold
    this.italic = opts.italic
    this.strikethrough = opts.strikethrough
    this.underline = opts.underline
    this.fontFamily = opts.fontFamily
    this.fontSize = opts.fontSize
    this.link = opts.link
    this.horizontalAlign = opts.horizontalAlign
    this.verticalAlign = opts.verticalAlign
    this.backgroundColor = opts.backgroundColor
    this.backgroundColorStyle = opts.backgroundColorStyle
    this.borders = opts.borders
    this.horizontalAlignment = opts.horizontalAlignment
    this.verticalAlignment = opts.verticalAlignment
    this.hyperlinkDisplayType = opts.hyperlinkDisplayType
    this.numberFormat = opts.numberFormat
    this.padding = opts.padding
    this.textDirection = opts.textDirection
    this.textFormat = opts.textFormat
    this.textRotation = opts.textRotation
    this.wrapStrategy = opts.wrapStrategy
    this.raw = opts.raw
  }

  toCellFormat(): CellFormat {
    if (this.raw) return this.raw

    const format: CellFormat = {}

    // backgroundColor — shorthand bgColor takes lowest precedence
    if (this.backgroundColor) {
      format.backgroundColor = this.backgroundColor
    } else if (this.bgColor) {
      const c = hexToColor(this.bgColor)
      if (c) format.backgroundColor = c
    }
    if (this.backgroundColorStyle) {
      format.backgroundColorStyle = this.backgroundColorStyle
    }

    // Borders
    if (this.borders) {
      const borders: Borders = {}
      if (this.borders.top) borders.top = toCellFormatBorder(this.borders.top)
      if (this.borders.bottom)
        borders.bottom = toCellFormatBorder(this.borders.bottom)
      if (this.borders.left)
        borders.left = toCellFormatBorder(this.borders.left)
      if (this.borders.right)
        borders.right = toCellFormatBorder(this.borders.right)
      if (Object.keys(borders).length > 0) format.borders = borders
    }

    // Horizontal alignment
    if (this.horizontalAlignment) {
      format.horizontalAlignment = this.horizontalAlignment
    } else if (this.horizontalAlign) {
      format.horizontalAlignment = HORIZONTAL_MAP[this.horizontalAlign]
    }

    // Vertical alignment
    if (this.verticalAlignment) {
      format.verticalAlignment = this.verticalAlignment
    } else if (this.verticalAlign) {
      format.verticalAlignment = VERTICAL_MAP[this.verticalAlign]
    }

    // Hyperlink display type
    if (this.hyperlinkDisplayType) {
      format.hyperlinkDisplayType = this.hyperlinkDisplayType
    }

    // Number format
    if (this.numberFormat) {
      format.numberFormat = this.numberFormat
    }

    // Padding
    if (this.padding) {
      format.padding = this.padding
    }

    // Text direction
    if (this.textDirection) {
      format.textDirection = this.textDirection
    }

    // Text format — merge explicit textFormat with shorthands
    const tf: TextFormat = { ...this.textFormat }

    if (this.underline !== undefined) tf.underline = this.underline
    if (this.bold !== undefined) tf.bold = this.bold
    if (this.italic !== undefined) tf.italic = this.italic
    if (this.strikethrough !== undefined) tf.strikethrough = this.strikethrough
    if (this.fontFamily !== undefined) tf.fontFamily = this.fontFamily
    if (this.fontSize !== undefined) tf.fontSize = this.fontSize
    if (this.link !== undefined) tf.link = this.link

    // textColor shorthand → foregroundColor (only if foregroundColor not already set)
    if (this.textColor && !tf.foregroundColor && !tf.foregroundColorStyle) {
      const c = hexToColor(this.textColor)
      if (c) tf.foregroundColor = c
    }

    if (Object.keys(tf).length > 0) format.textFormat = tf

    // Text rotation
    if (this.textRotation) {
      format.textRotation = this.textRotation
    }

    // Wrap strategy
    if (this.wrapStrategy) {
      format.wrapStrategy = this.wrapStrategy
    }

    return format
  }
}
