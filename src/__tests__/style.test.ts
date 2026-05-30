import { test, expect } from "bun:test"
import { Style } from "../style"

// Shorthand properties
test("Style stores and converts bgColor", () => {
  const s = new Style({ bgColor: "#ff0000" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 1, green: 0, blue: 0 })
})

test("Style stores and converts textColor", () => {
  const s = new Style({ textColor: "#00ff00" })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.foregroundColor).toEqual({ red: 0, green: 1, blue: 0 })
})

test("Style stores bold/italic/strikethrough", () => {
  const s = new Style({ bold: true, italic: true, strikethrough: true })
  const cf = s.toCellFormat()
  expect(cf.textFormat).toMatchObject({
    bold: true,
    italic: true,
    strikethrough: true,
  })
})

test("Style stores underline", () => {
  const s = new Style({ underline: true })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.underline).toBe(true)
})

test("Style stores fontFamily and fontSize", () => {
  const s = new Style({ fontFamily: "Arial", fontSize: 14 })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.fontFamily).toBe("Arial")
  expect(cf.textFormat?.fontSize).toBe(14)
})

test("Style stores horizontal alignment", () => {
  const s = new Style({ horizontalAlign: "center" })
  const cf = s.toCellFormat()
  expect(cf.horizontalAlignment).toBe("CENTER")
})

test("Style stores vertical alignment", () => {
  const s = new Style({ verticalAlign: "top" })
  const cf = s.toCellFormat()
  expect(cf.verticalAlignment).toBe("TOP")
})

// Full CellFormat properties
test("Style accepts raw backgroundColor object", () => {
  const s = new Style({ backgroundColor: { red: 0.5, green: 0.25, blue: 0 } })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 0.5, green: 0.25, blue: 0 })
})

test("Style backgroundColor overrides bgColor shorthand", () => {
  const s = new Style({
    bgColor: "#ff0000",
    backgroundColor: { red: 0, green: 0, blue: 1 },
  })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 0, green: 0, blue: 1 })
})

test("Style accepts backgroundColorStyle", () => {
  const s = new Style({
    backgroundColorStyle: { themeColor: "ACCENT1" },
  })
  const cf = s.toCellFormat()
  expect(cf.backgroundColorStyle).toEqual({ themeColor: "ACCENT1" })
})

test("Style accepts horizontalAlignment raw value", () => {
  const s = new Style({ horizontalAlignment: "RIGHT" })
  const cf = s.toCellFormat()
  expect(cf.horizontalAlignment).toBe("RIGHT")
})

test("Style horizontalAlignment overrides horizontalAlign shorthand", () => {
  const s = new Style({ horizontalAlign: "left", horizontalAlignment: "CENTER" })
  const cf = s.toCellFormat()
  expect(cf.horizontalAlignment).toBe("CENTER")
})

test("Style accepts verticalAlignment raw value", () => {
  const s = new Style({ verticalAlignment: "BOTTOM" })
  const cf = s.toCellFormat()
  expect(cf.verticalAlignment).toBe("BOTTOM")
})

// Borders
test("Style accepts borders with style shorthand", () => {
  const s = new Style({
    borders: {
      top: { style: "solid", color: { red: 0, green: 0, blue: 0 } },
      bottom: { style: "dashed" },
    },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.style).toBe("SOLID")
  expect(cf.borders?.top?.color).toEqual({ red: 0, green: 0, blue: 0 })
  expect(cf.borders?.bottom?.style).toBe("DASHED")
})

test("Style accepts borders with all shorthand styles", () => {
  const s = new Style({
    borders: {
      left: { style: "dotted" },
      right: { style: "double" },
    },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.left?.style).toBe("DOTTED")
  expect(cf.borders?.right?.style).toBe("DOUBLE")
})

test("Style handles border style none", () => {
  const s = new Style({
    borders: { top: { style: "none" } },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.style).toBe("NONE")
})

test("Style handles medium border styles", () => {
  const s = new Style({
    borders: {
      top: { style: "medium" },
      left: { style: "mediumDashed" },
      right: { style: "mediumDotted" },
      bottom: { style: "mediumSolid" },
    },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.style).toBe("MEDIUM")
  expect(cf.borders?.left?.style).toBe("MEDIUM_DASHED")
  expect(cf.borders?.right?.style).toBe("MEDIUM_DOTTED")
  expect(cf.borders?.bottom?.style).toBe("MEDIUM_SOLID")
})

test("Style handles thick border", () => {
  const s = new Style({
    borders: { top: { style: "thick" } },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.style).toBe("THICK")
})

test("Style handles border with colorStyle and width", () => {
  const s = new Style({
    borders: {
      top: { colorStyle: { themeColor: "TEXT" }, width: 3 },
    },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.colorStyle).toEqual({ themeColor: "TEXT" })
  expect(cf.borders?.top?.width).toBe(3)
})

test("Style omits borders when none provided", () => {
  const s = new Style({})
  const cf = s.toCellFormat()
  expect(cf.borders).toBeUndefined()
})

test("Style omits empty border sides", () => {
  const s = new Style({ borders: { top: {} } })
  const cf = s.toCellFormat()
  expect(cf.borders?.top).toBeUndefined()
})

// Hyperlink display type
test("Style accepts hyperlinkDisplayType", () => {
  const s = new Style({ hyperlinkDisplayType: "LINKED" })
  const cf = s.toCellFormat()
  expect(cf.hyperlinkDisplayType).toBe("LINKED")
})

test("Style accepts PLAIN_TEXT hyperlink display", () => {
  const s = new Style({ hyperlinkDisplayType: "PLAIN_TEXT" })
  const cf = s.toCellFormat()
  expect(cf.hyperlinkDisplayType).toBe("PLAIN_TEXT")
})

// Number format
test("Style accepts numberFormat", () => {
  const s = new Style({
    numberFormat: { type: "CURRENCY", pattern: "$#,##0.00" },
  })
  const cf = s.toCellFormat()
  expect(cf.numberFormat).toEqual({ type: "CURRENCY", pattern: "$#,##0.00" })
})

test("Style accepts numberFormat type only", () => {
  const s = new Style({ numberFormat: { type: "DATE" } })
  const cf = s.toCellFormat()
  expect(cf.numberFormat?.type).toBe("DATE")
})

// Padding
test("Style accepts padding", () => {
  const s = new Style({
    padding: { top: 5, bottom: 5, left: 10, right: 10 },
  })
  const cf = s.toCellFormat()
  expect(cf.padding).toEqual({ top: 5, bottom: 5, left: 10, right: 10 })
})

test("Style accepts partial padding", () => {
  const s = new Style({ padding: { left: 8 } })
  const cf = s.toCellFormat()
  expect(cf.padding?.left).toBe(8)
  expect(cf.padding?.top).toBeUndefined()
})

// Text direction
test("Style accepts textDirection", () => {
  const s = new Style({ textDirection: "RIGHT_TO_LEFT" })
  const cf = s.toCellFormat()
  expect(cf.textDirection).toBe("RIGHT_TO_LEFT")
})

// Text format passthrough
test("Style merges textFormat with shorthands", () => {
  const s = new Style({
    textFormat: { fontFamily: "Roboto" },
    bold: true,
    fontSize: 12,
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.fontFamily).toBe("Roboto")
  expect(cf.textFormat?.bold).toBe(true)
  expect(cf.textFormat?.fontSize).toBe(12)
})

test("Style textFormat foregroundColor takes precedence over textColor", () => {
  const s = new Style({
    textColor: "#ff0000",
    textFormat: { foregroundColor: { red: 0, green: 0, blue: 1 } },
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.foregroundColor).toEqual({ red: 0, green: 0, blue: 1 })
})

test("Style textFormat foregroundColorStyle takes precedence over textColor", () => {
  const s = new Style({
    textColor: "#ff0000",
    textFormat: { foregroundColorStyle: { themeColor: "TEXT" } },
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.foregroundColor).toBeUndefined()
  expect(cf.textFormat?.foregroundColorStyle).toEqual({ themeColor: "TEXT" })
})

test("Style link shorthand maps to textFormat.link", () => {
  const s = new Style({ link: { uri: "https://example.com" } })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.link).toEqual({ uri: "https://example.com" })
})

test("Style link merged with textFormat.link", () => {
  const s = new Style({
    link: { uri: "https://example.com" },
    bold: true,
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.link).toEqual({ uri: "https://example.com" })
  expect(cf.textFormat?.bold).toBe(true)
})

// Text rotation
test("Style accepts textRotation angle", () => {
  const s = new Style({ textRotation: { angle: 45 } })
  const cf = s.toCellFormat()
  expect(cf.textRotation?.angle).toBe(45)
})

test("Style accepts textRotation vertical", () => {
  const s = new Style({ textRotation: { vertical: true } })
  const cf = s.toCellFormat()
  expect(cf.textRotation?.vertical).toBe(true)
})

// Wrap strategy
test("Style accepts wrapStrategy", () => {
  const s = new Style({ wrapStrategy: "WRAP" })
  const cf = s.toCellFormat()
  expect(cf.wrapStrategy).toBe("WRAP")
})

test("Style accepts all wrap strategies", () => {
  for (const strat of ["OVERFLOW_CELL", "LEGACY_WRAP", "CLIP", "WRAP"] as const) {
    const s = new Style({ wrapStrategy: strat })
    expect(s.toCellFormat().wrapStrategy).toBe(strat)
  }
})

// Raw bypass
test("Style raw bypasses all other options", () => {
  const rawFormat = { backgroundColor: { red: 0.5, green: 0.5, blue: 0.5 } }
  const s = new Style({ raw: rawFormat, bold: true, bgColor: "#ff0000" })
  const cf = s.toCellFormat()
  expect(cf).toBe(rawFormat)
  expect(cf.backgroundColor).toEqual({ red: 0.5, green: 0.5, blue: 0.5 })
  expect(cf.textFormat).toBeUndefined()
})

// Edge cases
test("Style with no options returns empty format", () => {
  const s = new Style()
  const cf = s.toCellFormat()
  expect(cf).toEqual({})
})

test("Style invalid hex color does not throw", () => {
  const s = new Style({ bgColor: "not-a-hex" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toBeUndefined()
})

test("Style handles # prefix in hex", () => {
  const s = new Style({ bgColor: "#ffffff" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 1, green: 1, blue: 1 })
})

test("Style handles hex without # prefix", () => {
  const s = new Style({ bgColor: "000000" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 0, green: 0, blue: 0 })
})
