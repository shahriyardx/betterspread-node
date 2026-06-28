import { test, expect, mock } from "bun:test"
import { z } from "zod"
import { Cell } from "../cell"
import { ValidationError } from "../types"
import { Format } from "../format"
import type { sheets_v4 } from "@googleapis/sheets"

function makeMockTab(overrides: Record<string, unknown> = {}) {
  const mockGet = mock(() => Promise.resolve({ data: { values: [["42"]] } }))
  const mockUpdate = mock(() => Promise.resolve({ data: {} }))
  const mockClear = mock(() => Promise.resolve({ data: {} }))
  const mockBatchUpdate = mock(() => Promise.resolve({ data: {} }))

  const mockClient = {
    spreadsheets: {
      values: {
        get: mockGet,
        update: mockUpdate,
        clear: mockClear,
      },
      batchUpdate: mockBatchUpdate,
    },
  } as any as sheets_v4.Sheets

  const tab = {
    getClient: () => mockClient,
    getSheetId: () => "test-sheet-id",
    getTitle: () => "Sheet1",
    getWorksheetId: () => 0,
    getHeaders: () => [],
    getSchema: () => null,
    ...overrides,
  }

  return { tab, mockClient, mockGet, mockUpdate, mockClear, mockBatchUpdate }
}

test("Cell stores value and metadata", () => {
  const { tab } = makeMockTab()
  const cell = new Cell({
    value: "hello",
    label: "B",
    rowIndex: 3,
    cellIndex: 1,
    tab,
    row: null,
  })

  expect(cell.value).toBe("hello")
  expect(cell.label).toBe("B")
  expect(cell.rowIndex).toBe(3)
  expect(cell.cellIndex).toBe(1)
})

test("Cell toString returns value", () => {
  const { tab } = makeMockTab()
  const cell = new Cell({
    value: "test-value",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  expect(cell.toString()).toBe("test-value")
  expect(String(cell)).toBe("test-value")
})

test("Cell custom inspect shows row, header, and value", () => {
  const { tab } = makeMockTab()
  const cell = new Cell({
    value: "hello",
    label: "B",
    rowIndex: 3,
    cellIndex: 1,
    tab,
    row: null,
  })

  const inspected = cell[Symbol.for("nodejs.util.inspect.custom")]()
  expect(inspected).toBe('Cell(row=3, header="", value="hello")')
})

test("Cell header returns empty string when no headers cached", () => {
  const { tab } = makeMockTab()
  const cell = new Cell({
    value: "hello",
    label: "B",
    rowIndex: 3,
    cellIndex: 1,
    tab,
    row: null,
  })

  expect(cell.header).toBe("")
})

test("Cell header returns column header from tab cache", () => {
  const headers = ["Name", "Email", "Age"]
  const { tab } = makeMockTab({ getHeaders: () => headers })
  const cell = new Cell({
    value: "john@example.com",
    label: "B",
    rowIndex: 2,
    cellIndex: 1,
    tab,
    row: null,
  })

  expect(cell.header).toBe("Email")
})

test("Cell custom inspect shows header when headers cached", () => {
  const headers = ["Name", "Email", "Age"]
  const { tab } = makeMockTab({ getHeaders: () => headers })
  const cell = new Cell({
    value: "john@example.com",
    label: "B",
    rowIndex: 2,
    cellIndex: 1,
    tab,
    row: null,
  })

  const inspected = cell[Symbol.for("nodejs.util.inspect.custom")]()
  expect(inspected).toBe(
    'Cell(row=2, header="Email", value="john@example.com")',
  )
})

test("Cell stores optional format as Format wrapping CellFormat", () => {
  const { tab } = makeMockTab()
  const rawFormat = { backgroundColor: { red: 1, green: 0, blue: 0 } }
  const cell = new Cell({
    value: "x",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
    format: rawFormat,
  })

  expect(cell.format).toBeInstanceOf(Format)
  expect(cell.format!.backgroundColor).toEqual({ red: 1, green: 0, blue: 0 })
})

test("Cell format is null when not provided", () => {
  const { tab } = makeMockTab()
  const cell = new Cell({
    value: "x",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  expect(cell.format).toBeNull()
})

test("Cell update calls API and mutates in place", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "old",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const updated = await cell.update({ value: "new-value" })

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.updateCells.rows[0].values).toEqual([
    { userEnteredValue: { stringValue: "new-value" } },
  ])
  expect(updated.value).toBe("new-value")
  expect(updated).toBe(cell) // mutates and returns same instance
  expect(cell.value).toBe("new-value") // value updated in-place
})

test("Cell clear calls API", async () => {
  const { tab, mockClear } = makeMockTab()
  const cell = new Cell({
    value: "x",
    label: "B2",
    rowIndex: 2,
    cellIndex: 1,
    tab,
    row: null,
  })

  await cell.clear()
  expect(mockClear).toHaveBeenCalledTimes(1)
})

test("Cell style calls batchUpdate", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "x",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  await cell.style(
    new Format({ backgroundColor: { red: 1, green: 0, blue: 0 } }),
  )

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.repeatCell.range.startRowIndex).toBe(0)
  expect(req.repeatCell.range.endRowIndex).toBe(1)
  expect(req.repeatCell.range.startColumnIndex).toBe(0)
  expect(req.repeatCell.range.endColumnIndex).toBe(1)
})

test("Cell delete with shift=left sends COLUMNS", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "x",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  await cell.delete("left")

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteRange.shiftDimension).toBe("COLUMNS")
})

test("Cell delete with shift=up sends ROWS", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "x",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  await cell.delete("up")

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteRange.shiftDimension).toBe("ROWS")
})

test("Cell update passes through when no schema set", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "old",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const updated = await cell.update({ value: "new-value" })

  expect(updated.value).toBe("new-value")
  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})

test("Cell update validates against schema column type", async () => {
  const headers = ["Name", "Score"]
  const { tab } = makeMockTab({
    getHeaders: () => headers,
    getSchema: () => z.object({ Name: z.string(), Score: z.number() }),
  })
  const cell = new Cell({
    value: "10",
    label: "B",
    rowIndex: 1,
    cellIndex: 1,
    tab,
    row: null,
  })

  // header is "Score" which maps to z.number() — "abc" fails
  await expect(cell.update({ value: "abc" })).rejects.toThrow(ValidationError)
})

test("Cell update passes validation when schema matches", async () => {
  const headers = ["Name", "Score"]
  const { tab, mockBatchUpdate } = makeMockTab({
    getHeaders: () => headers,
    getSchema: () => z.object({ Name: z.string(), Score: z.string() }),
  })
  const cell = new Cell({
    value: "10",
    label: "B",
    rowIndex: 1,
    cellIndex: 1,
    tab,
    row: null,
  })

  const updated = await cell.update({ value: "99" })

  expect(updated.value).toBe("99")
  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})

test("Cell update with Cell object applies value and format via batchUpdate", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "old",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const newCell = new Cell({
    value: "new-val",
    format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
  })

  const updated = await cell.update({
    value: newCell.value,
    format: newCell.format ?? undefined,
  })

  expect(updated).toBe(cell)
  expect(cell.value).toBe("new-val")
  expect(cell.format).toBeInstanceOf(Format)
  expect(cell.format!.backgroundColor).toEqual({ red: 1, green: 0, blue: 0 })

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.updateCells.range.startRowIndex).toBe(0)
  expect(req.updateCells.range.endRowIndex).toBe(1)
  expect(req.updateCells.range.startColumnIndex).toBe(0)
  expect(req.updateCells.range.endColumnIndex).toBe(1)
  expect(req.updateCells.rows[0].values).toEqual([
    {
      userEnteredValue: { stringValue: "new-val" },
      userEnteredFormat: { backgroundColor: { red: 1, green: 0, blue: 0 } },
    },
  ])
  expect(req.updateCells.fields).toBe("userEnteredValue,userEnteredFormat")
})

test("Cell update with Cell object without format uses value only", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "old",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const newCell = new Cell({ value: "just-value" })
  await cell.update({ value: newCell.value })

  expect(cell.value).toBe("just-value")

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.updateCells.fields).toBe("userEnteredValue")
  expect(req.updateCells.rows[0].values).toEqual([
    { userEnteredValue: { stringValue: "just-value" } },
  ])
})

test("Cell delete defaults to shift=up (ROWS), matching Tab.delCell", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cell = new Cell({
    value: "x",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  await cell.delete()

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteRange.shiftDimension).toBe("ROWS")
})

test("Cell update coerces numeric string for number schema", async () => {
  const { tab, mockBatchUpdate } = makeMockTab({
    getHeaders: () => ["Name", "Score"],
    getSchema: () => z.object({ Name: z.string(), Score: z.number() }),
  })
  const cell = new Cell({
    value: "10",
    label: "B",
    rowIndex: 1,
    cellIndex: 1,
    tab,
    row: null,
  })

  const updated = await cell.update({ value: "42" })

  expect(updated.value).toBe("42")
  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})

test("Cell update rejects non-coercible string for number schema", async () => {
  const { tab } = makeMockTab({
    getHeaders: () => ["Name", "Score"],
    getSchema: () => z.object({ Name: z.string(), Score: z.number() }),
  })
  const cell = new Cell({
    value: "10",
    label: "B",
    rowIndex: 1,
    cellIndex: 1,
    tab,
    row: null,
  })

  await expect(cell.update({ value: "abc" })).rejects.toThrow(ValidationError)
})

test("Cell update coerces boolean string for boolean schema", async () => {
  const { tab, mockBatchUpdate } = makeMockTab({
    getHeaders: () => ["Active"],
    getSchema: () => z.object({ Active: z.boolean() }),
  })
  const cell = new Cell({
    value: "false",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const updated = await cell.update({ value: "true" })

  expect(updated.value).toBe("true")
  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})

test("Cell update coerces date string for date schema", async () => {
  const { tab, mockBatchUpdate } = makeMockTab({
    getHeaders: () => ["When"],
    getSchema: () => z.object({ When: z.date() }),
  })
  const cell = new Cell({
    value: "",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const updated = await cell.update({ value: "2026-06-28" })

  expect(updated.value).toBe("2026-06-28")
  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})
