import { test, expect } from "bun:test"
import { columnLabel, columnIndex, parseCellAddress, hexToColor } from "../utils"

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

test("hexToColor returns null for invalid hex", () => {
  expect(hexToColor("xyz")).toBeNull()
})
