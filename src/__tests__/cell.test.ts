import { test, expect, mock } from "bun:test"
import { Cell, type CellUpdateOptions } from "../cell"
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
  expect(inspected).toBe(
    'Cell(row=3, header="", value="hello")',
  )
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

test("Cell update calls API and returns new cell", async () => {
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
  expect(updated).not.toBe(cell)
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
