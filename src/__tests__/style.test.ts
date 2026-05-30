import { test, expect, mock } from "bun:test"
import { Style } from "../style"

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
  expect(cf.textFormat).toEqual({
    bold: true,
    italic: true,
    strikethrough: true,
  })
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

test("Style raw bypasses all other options", () => {
  const rawFormat = { backgroundColor: { red: 0.5, green: 0.5, blue: 0.5 } }
  const s = new Style({ raw: rawFormat, bold: true, bgColor: "#ff0000" })
  const cf = s.toCellFormat()
  expect(cf).toBe(rawFormat)
  expect(cf.backgroundColor).toEqual({ red: 0.5, green: 0.5, blue: 0.5 })
  expect(cf.textFormat).toBeUndefined()
})

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
