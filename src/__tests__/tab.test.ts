import { test, expect, mock } from "bun:test"
import { z } from "zod"
import { Tab } from "../tab"
import { Cell } from "../cell"
import { ValidationError } from "../types"
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
      data: {
        sheets: [
          {
            properties: { title: "Sheet1", sheetId: 0 },
            data: [
              {
                rowData: [
                  {
                    values: [
                      { effectiveValue: { stringValue: "a" } },
                      { effectiveValue: { stringValue: "b" } },
                    ],
                  },
                  {
                    values: [
                      { effectiveValue: { stringValue: "c" } },
                      { effectiveValue: { stringValue: "d" } },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
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

  return {
    sheet,
    mockClient,
    mockGet,
    mockAppend,
    mockBatchUpdate,
    mockGetSpreadsheet,
  }
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
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.values({ range: "Sheet1!A1:B2" })

  expect(mockGetSpreadsheet.mock.calls[0]?.[0].ranges?.[0]).toBe("Sheet1!A1:B2")
})

test("Tab values with format uses spreadsheets.get and returns format", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  mockGetSpreadsheet.mockImplementation(() =>
    Promise.resolve({
      data: {
        sheets: [
          {
            data: [
              {
                rowData: [
                  {
                    values: [
                      {
                        effectiveValue: { stringValue: "a" },
                        userEnteredFormat: {
                          backgroundColor: { red: 1, green: 0, blue: 0 },
                        },
                      },
                      { effectiveValue: { stringValue: "b" } },
                    ],
                  },
                  {
                    values: [
                      { effectiveValue: { stringValue: "c" } },
                      { effectiveValue: { numberValue: 42 } },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    }),
  )
  const tab = new Tab(sheet, "Sheet1", 0, {})
  const rows = await tab.values()

  expect(rows).toHaveLength(2)
  expect(rows[0][0].value).toBe("a")
  expect(rows[0][1].value).toBe("b")
  expect(rows[1][0].value).toBe("c")
  expect(rows[1][1].value).toBe("42")

  // Format populated
  expect(rows[0][0].format?.backgroundColor).toEqual({
    red: 1,
    green: 0,
    blue: 0,
  })
  // Cell without format
  expect(rows[0][1].format).toBeNull()

  // Uses spreadsheets.get with correct fields
  const callArgs = mockGetSpreadsheet.mock.calls[0]?.[0]
  expect(callArgs.fields).toContain("userEnteredFormat")
})

test("Tab values format uses NUMBER value type", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  mockGetSpreadsheet.mockImplementation(() =>
    Promise.resolve({
      data: {
        sheets: [
          {
            data: [
              {
                rowData: [
                  {
                    values: [
                      { effectiveValue: { numberValue: 100 } },
                      { effectiveValue: { boolValue: true } },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    }),
  )
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const rows = await tab.values()

  expect(rows[0][0].value).toBe("100")
  expect(rows[0][1].value).toBe("true")
})

test("Tab values with FORMULA render option uses spreadsheets.values.get", async () => {
  const { sheet, mockGet } = makeMockSheet()
  mockGet.mockImplementation(() =>
    Promise.resolve({
      data: {
        values: [["=SUM(A1:A10)", "plain"]],
      },
    }),
  )
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const rows = await tab.values({
    range: "Sheet1!A1:B1",
    valueRenderOption: "FORMULA",
  })

  expect(mockGet.mock.calls[0]?.[0].valueRenderOption).toBe("FORMULA")
  expect(rows[0][0].value).toBe("=SUM(A1:A10)")
  expect(rows[0][1].value).toBe("plain")
  // No format data in this path
  expect(rows[0][0].format).toBeNull()
  expect(rows[0][1].format).toBeNull()
})

test("Tab getRow with FORMULA render option passes through", async () => {
  const { sheet, mockGet } = makeMockSheet()
  mockGet.mockImplementation(() =>
    Promise.resolve({
      data: {
        values: [["=SUM(A1:A10)"]],
      },
    }),
  )
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const row = await tab.getRow({ serialNo: 1, valueRenderOption: "FORMULA" })

  expect(mockGet.mock.calls[0]?.[0].valueRenderOption).toBe("FORMULA")
  expect(row[0].value).toBe("=SUM(A1:A10)")
})

test("Tab getRow returns Row with correct index", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const row = await tab.getRow({ serialNo: 1 })

  expect(mockGetSpreadsheet.mock.calls[0]?.[0].ranges?.[0]).toContain("A1")
  expect(row).toHaveLength(2)
  expect(row[0].value).toBe("a")
})

test("Tab getRow with empty result returns empty Row", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  mockGetSpreadsheet.mockImplementation(() =>
    Promise.resolve({ data: { sheets: [{ data: [{ rowData: [] }] }] } }),
  )
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const row = await tab.getRow({ serialNo: 5 })
  expect(row).toHaveLength(0)
})

test("Tab getCell returns Cell with correct metadata", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const cell = await tab.getCell({ cellName: "B3" })

  expect(cell.value).toBe("a")
  expect(cell.label).toBe("B")
  expect(cell.rowIndex).toBe(3)
  expect(cell.cellIndex).toBe(1)
})

test("Tab getCell passes valueRenderOption to API", async () => {
  const { sheet, mockGet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.getCell({ cellName: "A1", valueRenderOption: "FORMULA" })

  expect(mockGet.mock.calls[0]?.[0].valueRenderOption).toBe("FORMULA")
})

test("Tab append without getRow returns null", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const result = await tab.append({ values: ["x", "y"] })

  expect(result).toBeNull()
  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.appendCells.rows[0].values).toEqual([
    { userEnteredValue: { stringValue: "x" } },
    { userEnteredValue: { stringValue: "y" } },
  ])
})

test("Tab append with Cell objects extracts values and applies format", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const styledCell = new Cell({
    value: "x",
    format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
  })

  await tab.append({ values: [styledCell, "y"] })

  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.appendCells.rows[0].values).toEqual([
    {
      userEnteredValue: { stringValue: "x" },
      userEnteredFormat: { backgroundColor: { red: 1, green: 0, blue: 0 } },
    },
    { userEnteredValue: { stringValue: "y" } },
  ])
})

test("Tab append with getRow returns Row", async () => {
  const { sheet, mockGet } = makeMockSheet()
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

  const result = await tab.append({ values: ["x", "y"], getRow: true })

  expect(result).not.toBeNull()
  expect(result).toHaveLength(2)
  expect(result?.[0].value).toBe("x")
  expect(result?.[1].value).toBe("y")
})

test("Tab delRow calls batchUpdate with correct range", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delRow({ start: 3 })

  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.deleteDimension.range.dimension).toBe("ROWS")
  expect(req.deleteDimension.range.startIndex).toBe(2)
  expect(req.deleteDimension.range.endIndex).toBe(3)
})

test("Tab delRow with end range", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delRow({ start: 3, end: 5 })

  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.deleteDimension.range.startIndex).toBe(2)
  expect(req.deleteDimension.range.endIndex).toBe(5)
})

test("Tab delCell calls batchUpdate with deleteRange", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delCell({ start: "B2" })

  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.deleteRange.range.sheetId).toBe(0)
  expect(req.deleteRange.shiftDimension).toBe("ROWS")
})

test("Tab delCell with range and shift left", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.delCell({ start: "A1", end: "C3", shift: "left" })

  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.deleteRange.range.startRowIndex).toBe(0)
  expect(req.deleteRange.range.endRowIndex).toBe(3)
  expect(req.deleteRange.range.startColumnIndex).toBe(0)
  expect(req.deleteRange.range.endColumnIndex).toBe(3)
  expect(req.deleteRange.shiftDimension).toBe("COLUMNS")
})

test("Tab delCell throws on invalid address", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  expect(tab.delCell({ start: "invalid" })).rejects.toThrow(
    /Invalid cell address/,
  )
})

test("Tab setSchema stores schema and returns this", () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  const schema = z.object({ Name: z.string(), Age: z.string() })

  const result = tab.setSchema(schema)

  expect(result).toBe(tab)
  expect(tab.getSchema()).toBe(schema)
})

test("Tab append with object validates and converts to row", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Email", "Age"])

  await tab.append({
    values: { Name: "John", Email: "john@test.com", Age: "30" },
  })

  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.appendCells.rows[0].values).toEqual([
    { userEnteredValue: { stringValue: "John" } },
    { userEnteredValue: { stringValue: "john@test.com" } },
    { userEnteredValue: { stringValue: "30" } },
  ])
})

test("Tab append with object validates against schema", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Email", "Age"])
  tab.setSchema(
    z.object({ Name: z.string(), Email: z.string(), Age: z.string() }),
  )

  await tab.append({
    values: { Name: "John", Email: "john@test.com", Age: "30" },
  })

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})

test("Tab append with object throws ValidationError on invalid schema", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Age"])
  tab.setSchema(z.object({ Name: z.string(), Age: z.number() }))

  expect(
    tab.append({ values: { Name: "John", Age: "not-a-number" } }),
  ).rejects.toThrow(ValidationError)
})

test("Tab append with Cell[] extracts values", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const cell1 = new Cell({ value: "x" })

  await tab.append({ values: [cell1, "y"] })

  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.appendCells.rows[0].values).toEqual([
    { userEnteredValue: { stringValue: "x" } },
    { userEnteredValue: { stringValue: "y" } },
  ])
})

test("Tab append array validates against schema by column", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Score"])
  tab.setSchema(z.object({ name: z.string(), score: z.number() }))

  // "abc" fails z.number() for Score column
  expect(tab.append({ values: ["A", "abc"] })).rejects.toThrow(ValidationError)
})

test("Tab append array passes valid values", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Score"])
  tab.setSchema(z.object({ name: z.string(), score: z.string() }))

  await tab.append({ values: ["Alice", "95"] })

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})
