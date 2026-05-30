import { test, expect, mock } from "bun:test"
import { Tab } from "../tab"
import type { sheets_v4 } from "@googleapis/sheets"

function makeMockSheet() {
  const mockGet = mock(() =>
    Promise.resolve({
      data: {
        values: [
          ["a", "b"],
          ["c", "d"],
        ],
      },
    }),
  )
  const mockAppend = mock(() => Promise.resolve({ data: { updates: {} } }))
  const mockBatchUpdate = mock(() => Promise.resolve({ data: {} }))
  const mockGetSpreadsheet = mock(() =>
    Promise.resolve({
      data: { sheets: [{ properties: { title: "Sheet1", sheetId: 0 } }] },
    }),
  )

  const mockClient = {
    spreadsheets: {
      values: {
        get: mockGet,
        append: mockAppend,
      },
      batchUpdate: mockBatchUpdate,
      get: mockGetSpreadsheet,
    },
  } as any as sheets_v4.Sheets

  const sheet = {
    getClient: () => mockClient,
    getId: () => "test-sheet-id",
    getTitle: () => "Sheet1",
  }

  return { sheet, mockClient, mockGet, mockAppend, mockBatchUpdate }
}

test("Tab values returns array of Rows", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const rows = await tab.values()

  expect(rows).toHaveLength(2)
  expect(rows[0][0].value).toBe("a")
  expect(rows[1][1].value).toBe("d")
})

test("Tab values with custom range", async () => {
  const { sheet, mockGet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.values("Sheet1!A1:B2")

  expect(mockGet.mock.calls[0][0].range).toBe("Sheet1!A1:B2")
})

test("Tab getRow returns Row with correct index", async () => {
  const { sheet, mockGet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const row = await tab.getRow(1)

  expect(mockGet.mock.calls[0][0].range).toContain("A1")
  expect(row).toHaveLength(2)
  expect(row[0].value).toBe("a")
})

test("Tab getRow with empty result returns empty Row", async () => {
  const { sheet, mockGet } = makeMockSheet()
  mockGet.mockImplementation(() => Promise.resolve({ data: { values: [] } }))
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const row = await tab.getRow(5)
  expect(row).toHaveLength(0)
})

test("Tab getCell returns Cell with correct metadata", async () => {
  const { sheet, mockGet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const cell = await tab.getCell("B3")

  expect(cell.value).toBe("a")
  expect(cell.label).toBe("B")
  expect(cell.rowIndex).toBe(3)
  expect(cell.cellIndex).toBe(1)
})

test("Tab getCell passes render_option to API", async () => {
  const { sheet, mockGet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.getCell("A1", "formula")

  expect(mockGet.mock.calls[0][0].valueRenderOption).toBe("FORMULA")
})

test("Tab append without getRow returns null", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const result = await tab.append(["x", "y"])

  expect(result).toBeNull()
  expect(mockAppend).toHaveBeenCalledTimes(1)
  expect(mockAppend.mock.calls[0][0].requestBody.values).toEqual([["x", "y"]])
})

test("Tab append with getRow returns Row", async () => {
  const { sheet, mockAppend, mockGet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  mockGet.mockImplementation(() =>
    Promise.resolve({
      data: {
        values: [
          ["a", "b"],
          ["c", "d"],
          ["x", "y"],
        ],
      },
    }),
  )

  const result = await tab.append(["x", "y"], true)

  expect(result).not.toBeNull()
  expect(result).toHaveLength(2)
  expect(result![0].value).toBe("x")
  expect(result![1].value).toBe("y")
})

test("Tab delRow calls batchUpdate with correct range", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delRow(3)

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteDimension.range.dimension).toBe("ROWS")
  expect(req.deleteDimension.range.startIndex).toBe(2)
  expect(req.deleteDimension.range.endIndex).toBe(3)
})

test("Tab delRow with end range", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delRow(3, 5)

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteDimension.range.startIndex).toBe(2)
  expect(req.deleteDimension.range.endIndex).toBe(5)
})

test("Tab delCell calls batchUpdate with deleteRange", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delCell("B2")

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteRange.range.sheetId).toBe(0)
  expect(req.deleteRange.shiftDimension).toBe("ROWS")
})

test("Tab delCell with range and shift left", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delCell("A1", "C3", "left")

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteRange.range.startRowIndex).toBe(0)
  expect(req.deleteRange.range.endRowIndex).toBe(3)
  expect(req.deleteRange.range.startColumnIndex).toBe(0)
  expect(req.deleteRange.range.endColumnIndex).toBe(3)
  expect(req.deleteRange.shiftDimension).toBe("COLUMNS")
})

test("Tab delCell throws on invalid address", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  expect(tab.delCell("invalid")).rejects.toThrow(/Invalid cell address/)
})
