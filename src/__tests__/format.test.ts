import { test, expect } from "bun:test"
import { Format } from "../format"

// Shorthand properties
test("Format stores and converts bgColor", () => {
  const s = new Format({ bgColor: "#ff0000" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 1, green: 0, blue: 0 })
})

test("Format stores and converts textColor", () => {
  const s = new Format({ textColor: "#00ff00" })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.foregroundColor).toEqual({ red: 0, green: 1, blue: 0 })
})

test("Format stores bold/italic/strikethrough", () => {
  const s = new Format({ bold: true, italic: true, strikethrough: true })
  const cf = s.toCellFormat()
  expect(cf.textFormat).toMatchObject({
    bold: true,
    italic: true,
    strikethrough: true,
  })
})

test("Format stores underline", () => {
  const s = new Format({ underline: true })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.underline).toBe(true)
})

test("Format stores fontFamily and fontSize", () => {
  const s = new Format({ fontFamily: "Arial", fontSize: 14 })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.fontFamily).toBe("Arial")
  expect(cf.textFormat?.fontSize).toBe(14)
})

test("Format stores horizontal alignment", () => {
  const s = new Format({ horizontalAlign: "center" })
  const cf = s.toCellFormat()
  expect(cf.horizontalAlignment).toBe("CENTER")
})

test("Format stores vertical alignment", () => {
  const s = new Format({ verticalAlign: "top" })
  const cf = s.toCellFormat()
  expect(cf.verticalAlignment).toBe("TOP")
})

// Full CellFormat properties
test("Format accepts raw backgroundColor object", () => {
  const s = new Format({ backgroundColor: { red: 0.5, green: 0.25, blue: 0 } })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 0.5, green: 0.25, blue: 0 })
})

test("Format backgroundColor overrides bgColor shorthand", () => {
  const s = new Format({
    bgColor: "#ff0000",
    backgroundColor: { red: 0, green: 0, blue: 1 },
  })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 0, green: 0, blue: 1 })
})

test("Format accepts backgroundColorStyle", () => {
  const s = new Format({
    backgroundColorStyle: { themeColor: "ACCENT1" },
  })
  const cf = s.toCellFormat()
  expect(cf.backgroundColorStyle).toEqual({ themeColor: "ACCENT1" })
})

test("Format accepts horizontalAlignment raw value", () => {
  const s = new Format({ horizontalAlignment: "RIGHT" })
  const cf = s.toCellFormat()
  expect(cf.horizontalAlignment).toBe("RIGHT")
})

test("Format horizontalAlignment overrides horizontalAlign shorthand", () => {
  const s = new Format({ horizontalAlign: "left", horizontalAlignment: "CENTER" })
  const cf = s.toCellFormat()
  expect(cf.horizontalAlignment).toBe("CENTER")
})

test("Format accepts verticalAlignment raw value", () => {
  const s = new Format({ verticalAlignment: "BOTTOM" })
  const cf = s.toCellFormat()
  expect(cf.verticalAlignment).toBe("BOTTOM")
})

// Borders
test("Format accepts borders with style shorthand", () => {
  const s = new Format({
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

test("Format accepts borders with all shorthand styles", () => {
  const s = new Format({
    borders: {
      left: { style: "dotted" },
      right: { style: "double" },
    },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.left?.style).toBe("DOTTED")
  expect(cf.borders?.right?.style).toBe("DOUBLE")
})

test("Format handles border style none", () => {
  const s = new Format({
    borders: { top: { style: "none" } },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.style).toBe("NONE")
})

test("Format handles medium border styles", () => {
  const s = new Format({
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

test("Format handles thick border", () => {
  const s = new Format({
    borders: { top: { style: "thick" } },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.style).toBe("THICK")
})

test("Format handles border with colorStyle and width", () => {
  const s = new Format({
    borders: {
      top: { colorStyle: { themeColor: "TEXT" }, width: 3 },
    },
  })
  const cf = s.toCellFormat()
  expect(cf.borders?.top?.colorStyle).toEqual({ themeColor: "TEXT" })
  expect(cf.borders?.top?.width).toBe(3)
})

test("Format omits borders when none provided", () => {
  const s = new Format({})
  const cf = s.toCellFormat()
  expect(cf.borders).toBeUndefined()
})

test("Format omits empty border sides", () => {
  const s = new Format({ borders: { top: {} } })
  const cf = s.toCellFormat()
  expect(cf.borders?.top).toBeUndefined()
})

// Hyperlink display type
test("Format accepts hyperlinkDisplayType", () => {
  const s = new Format({ hyperlinkDisplayType: "LINKED" })
  const cf = s.toCellFormat()
  expect(cf.hyperlinkDisplayType).toBe("LINKED")
})

test("Format accepts PLAIN_TEXT hyperlink display", () => {
  const s = new Format({ hyperlinkDisplayType: "PLAIN_TEXT" })
  const cf = s.toCellFormat()
  expect(cf.hyperlinkDisplayType).toBe("PLAIN_TEXT")
})

// Number format
test("Format accepts numberFormat", () => {
  const s = new Format({
    numberFormat: { type: "CURRENCY", pattern: "$#,##0.00" },
  })
  const cf = s.toCellFormat()
  expect(cf.numberFormat).toEqual({ type: "CURRENCY", pattern: "$#,##0.00" })
})

test("Format accepts numberFormat type only", () => {
  const s = new Format({ numberFormat: { type: "DATE" } })
  const cf = s.toCellFormat()
  expect(cf.numberFormat?.type).toBe("DATE")
})

// Padding
test("Format accepts padding", () => {
  const s = new Format({
    padding: { top: 5, bottom: 5, left: 10, right: 10 },
  })
  const cf = s.toCellFormat()
  expect(cf.padding).toEqual({ top: 5, bottom: 5, left: 10, right: 10 })
})

test("Format accepts partial padding", () => {
  const s = new Format({ padding: { left: 8 } })
  const cf = s.toCellFormat()
  expect(cf.padding?.left).toBe(8)
  expect(cf.padding?.top).toBeUndefined()
})

// Text direction
test("Format accepts textDirection", () => {
  const s = new Format({ textDirection: "RIGHT_TO_LEFT" })
  const cf = s.toCellFormat()
  expect(cf.textDirection).toBe("RIGHT_TO_LEFT")
})

// Text format passthrough
test("Format merges textFormat with shorthands", () => {
  const s = new Format({
    textFormat: { fontFamily: "Roboto" },
    bold: true,
    fontSize: 12,
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.fontFamily).toBe("Roboto")
  expect(cf.textFormat?.bold).toBe(true)
  expect(cf.textFormat?.fontSize).toBe(12)
})

test("Format textFormat foregroundColor takes precedence over textColor", () => {
  const s = new Format({
    textColor: "#ff0000",
    textFormat: { foregroundColor: { red: 0, green: 0, blue: 1 } },
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.foregroundColor).toEqual({ red: 0, green: 0, blue: 1 })
})

test("Format textFormat foregroundColorStyle takes precedence over textColor", () => {
  const s = new Format({
    textColor: "#ff0000",
    textFormat: { foregroundColorStyle: { themeColor: "TEXT" } },
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.foregroundColor).toBeUndefined()
  expect(cf.textFormat?.foregroundColorStyle).toEqual({ themeColor: "TEXT" })
})

test("Format link shorthand maps to textFormat.link", () => {
  const s = new Format({ link: { uri: "https://example.com" } })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.link).toEqual({ uri: "https://example.com" })
})

test("Format link merged with textFormat.link", () => {
  const s = new Format({
    link: { uri: "https://example.com" },
    bold: true,
  })
  const cf = s.toCellFormat()
  expect(cf.textFormat?.link).toEqual({ uri: "https://example.com" })
  expect(cf.textFormat?.bold).toBe(true)
})

// Text rotation
test("Format accepts textRotation angle", () => {
  const s = new Format({ textRotation: { angle: 45 } })
  const cf = s.toCellFormat()
  expect(cf.textRotation?.angle).toBe(45)
})

test("Format accepts textRotation vertical", () => {
  const s = new Format({ textRotation: { vertical: true } })
  const cf = s.toCellFormat()
  expect(cf.textRotation?.vertical).toBe(true)
})

// Wrap strategy
test("Format accepts wrapStrategy", () => {
  const s = new Format({ wrapStrategy: "WRAP" })
  const cf = s.toCellFormat()
  expect(cf.wrapStrategy).toBe("WRAP")
})

test("Format accepts all wrap strategies", () => {
  for (const strat of ["OVERFLOW_CELL", "LEGACY_WRAP", "CLIP", "WRAP"] as const) {
    const s = new Format({ wrapStrategy: strat })
    expect(s.toCellFormat().wrapStrategy).toBe(strat)
  }
})

// Raw bypass
test("Format raw bypasses all other options", () => {
  const rawFormat = { backgroundColor: { red: 0.5, green: 0.5, blue: 0.5 } }
  const s = new Format({ raw: rawFormat, bold: true, bgColor: "#ff0000" })
  const cf = s.toCellFormat()
  expect(cf).toBe(rawFormat)
  expect(cf.backgroundColor).toEqual({ red: 0.5, green: 0.5, blue: 0.5 })
  expect(cf.textFormat).toBeUndefined()
})

// Edge cases
test("Format with no options returns empty format", () => {
  const s = new Format()
  const cf = s.toCellFormat()
  expect(cf).toEqual({})
})

test("Format invalid hex color does not throw", () => {
  const s = new Format({ bgColor: "not-a-hex" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toBeUndefined()
})

test("Format handles # prefix in hex", () => {
  const s = new Format({ bgColor: "#ffffff" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 1, green: 1, blue: 1 })
})

test("Format handles hex without # prefix", () => {
  const s = new Format({ bgColor: "000000" })
  const cf = s.toCellFormat()
  expect(cf.backgroundColor).toEqual({ red: 0, green: 0, blue: 0 })
})
