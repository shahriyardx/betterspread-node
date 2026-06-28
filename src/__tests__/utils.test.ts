import { test, expect } from "bun:test"
import {
  columnLabel,
  columnIndex,
  parseCellAddress,
  hexToColor,
} from "../utils"

test("columnLabel converts 0 to A", () => {
  expect(columnLabel(0)).toBe("A")
})

test("columnLabel converts 25 to Z", () => {
  expect(columnLabel(25)).toBe("Z")
})

test("columnLabel converts 26 to AA", () => {
  expect(columnLabel(26)).toBe("AA")
})

test("columnLabel converts 51 to AZ", () => {
  expect(columnLabel(51)).toBe("AZ")
})

test("columnLabel converts 701 to ZZ", () => {
  expect(columnLabel(701)).toBe("ZZ")
})

test("columnIndex converts A to 0", () => {
  expect(columnIndex("A")).toBe(0)
})

test("columnIndex converts Z to 25", () => {
  expect(columnIndex("Z")).toBe(25)
})

test("columnIndex converts AA to 26", () => {
  expect(columnIndex("AA")).toBe(26)
})

test("columnIndex converts AZ to 51", () => {
  expect(columnIndex("AZ")).toBe(51)
})

test("columnIndex converts ZZ to 701", () => {
  expect(columnIndex("ZZ")).toBe(701)
})

test("columnLabel and columnIndex are inverses", () => {
  for (const n of [0, 1, 25, 26, 51, 100, 701, 702, 1000]) {
    expect(columnIndex(columnLabel(n))).toBe(n)
  }
})

test("parseCellAddress returns label and row", () => {
  const addr = parseCellAddress("B2")
  expect(addr).toEqual({ label: "B", row: 2 })
})

test("parseCellAddress handles multi-char column", () => {
  const addr = parseCellAddress("AA10")
  expect(addr).toEqual({ label: "AA", row: 10 })
})

test("parseCellAddress returns null for invalid address", () => {
  expect(parseCellAddress("invalid")).toBeNull()
})

test("parseCellAddress returns null for empty string", () => {
  expect(parseCellAddress("")).toBeNull()
})

test("hexToColor converts #ff0000 to RGB", () => {
  expect(hexToColor("#ff0000")).toEqual({ red: 1, green: 0, blue: 0 })
})

test("hexToColor converts #00ff00 to RGB", () => {
  expect(hexToColor("#00ff00")).toEqual({ red: 0, green: 1, blue: 0 })
})

test("hexToColor converts #0000ff to RGB", () => {
  expect(hexToColor("#0000ff")).toEqual({ red: 0, green: 0, blue: 1 })
})

test("hexToColor handles without #", () => {
  expect(hexToColor("ffffff")).toEqual({ red: 1, green: 1, blue: 1 })
})

test("hexToColor handles case insensitive", () => {
  expect(hexToColor("#FFAABB")).toEqual({
    red: 1,
    green: 0.6666666666666666,
    blue: 0.7333333333333333,
  })
})

test("hexToColor throws on invalid hex string", () => {
  expect(() => hexToColor("xyz")).toThrow(/Invalid hex color/)
})

test("hexToColor expands 3-digit shorthand #F0F", () => {
  expect(hexToColor("#F0F")).toEqual(hexToColor("#FF00FF"))
})

test("hexToColor expands 3-digit shorthand without #", () => {
  expect(hexToColor("F0F")).toEqual({ red: 1, green: 0, blue: 1 })
})

test("hexToColor expands mixed case #a1b", () => {
  expect(hexToColor("#a1b")).toEqual({
    red: parseInt("aa", 16) / 255,
    green: parseInt("11", 16) / 255,
    blue: parseInt("bb", 16) / 255,
  })
})

test("hexToColor throws on wrong length", () => {
  expect(() => hexToColor("FF00")).toThrow(/Invalid hex color/)
})

test("hexToColor throws on non-hex chars", () => {
  expect(() => hexToColor("#GGGGGG")).toThrow(/Invalid hex color/)
})

test("columnIndex normalizes lowercase input", () => {
  expect(columnIndex("a")).toBe(0)
  expect(columnIndex("A")).toBe(0)
  expect(columnIndex("z")).toBe(25)
  expect(columnIndex("aa")).toBe(26)
})

test("columnIndex returns -1 for invalid input", () => {
  expect(columnIndex("123")).toBe(-1)
  expect(columnIndex("A1")).toBe(-1)
  expect(columnIndex("")).toBe(-1)
})
