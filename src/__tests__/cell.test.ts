import { test, expect, mock } from "bun:test"
import { z } from "zod"
import { Cell, type CellUpdateOptions } from "../cell"
import { ValidationError } from "../types"
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

test("Cell stores optional format from CellFormat object", () => {
  const { tab } = makeMockTab()
  const format = { backgroundColor: { red: 1, green: 0, blue: 0 } }
  const cell = new Cell({
    value: "x",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
    format,
  })

  expect(cell.format).toEqual(format)
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
  const { tab, mockUpdate } = makeMockTab()
  const cell = new Cell({
    value: "old",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const updated = await cell.update("new-value")

  expect(mockUpdate).toHaveBeenCalledTimes(1)
  expect(mockUpdate.mock.calls[0][0].requestBody.values).toEqual([
    ["new-value"],
  ])
  expect(updated.value).toBe("new-value")
  expect(updated).toBe(cell) // mutates and returns same instance
  expect(cell.value).toBe("new-value") // value updated in-place
})

test("Cell update with formula input_format", async () => {
  const { tab, mockUpdate } = makeMockTab()
  const cell = new Cell({
    value: "",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  await cell.update("=SUM(A1:A10)", { inputFormat: "user_entered" })

  expect(mockUpdate.mock.calls[0][0].valueInputOption).toBe("USER_ENTERED")
})

test("Cell update with render_format triggers refetch", async () => {
  const { tab, mockGet, mockUpdate } = makeMockTab()
  const cell = new Cell({
    value: "",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  await cell.update("=SUM(A1:A10)", { renderFormat: "formula" })

  expect(mockGet).toHaveBeenCalledTimes(1)
  expect(mockGet.mock.calls[0][0].valueRenderOption).toBe("FORMULA")
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

  await cell.style({ backgroundColor: { red: 1, green: 0, blue: 0 } })

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
  const { tab, mockUpdate } = makeMockTab()
  const cell = new Cell({
    value: "old",
    label: "A",
    rowIndex: 1,
    cellIndex: 0,
    tab,
    row: null,
  })

  const updated = await cell.update("new-value")

  expect(updated.value).toBe("new-value")
  expect(mockUpdate).toHaveBeenCalledTimes(1)
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
  await expect(cell.update("abc")).rejects.toThrow(ValidationError)
})

test("Cell update passes validation when schema matches", async () => {
  const headers = ["Name", "Score"]
  const { tab, mockUpdate } = makeMockTab({
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

  const updated = await cell.update("99")

  expect(updated.value).toBe("99")
  expect(mockUpdate).toHaveBeenCalledTimes(1)
})
