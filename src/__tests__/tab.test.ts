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
  const mockAppend = mock(() =>
    Promise.resolve({
      data: { updates: { updatedRange: "Sheet1!A3:B3" } },
    }),
  )
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

test("Tab append without getRow returns null and writes via values.append", async () => {
  const { sheet, mockAppend, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const result = await tab.append({ values: ["x", "y"] })

  expect(result).toBeNull()
  // Plain values go through values.append, not appendCells/batchUpdate.
  expect(mockAppend).toHaveBeenCalledTimes(1)
  const appendReq = mockAppend.mock.calls[0]?.[0]
  expect(appendReq.insertDataOption).toBe("INSERT_ROWS")
  expect(appendReq.valueInputOption).toBe("RAW")
  expect(appendReq.requestBody.values).toEqual([["x", "y"]])
  // No format → no second format pass.
  expect(mockBatchUpdate).not.toHaveBeenCalled()
})

test("Tab append with Cell objects writes values then applies format", async () => {
  const { sheet, mockAppend, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const styledCell = new Cell({
    value: "x",
    format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
  })

  await tab.append({ values: [styledCell, "y"] })

  // Values written via values.append.
  expect(mockAppend.mock.calls[0]?.[0].requestBody.values).toEqual([["x", "y"]])
  // Format applied in a follow-up updateCells over the echoed row (A3:B3 → row 3).
  const req = mockBatchUpdate.mock.calls[0]?.[0].requestBody.requests[0]
  expect(req.updateCells.range).toEqual({
    sheetId: 0,
    startRowIndex: 2,
    endRowIndex: 3,
    startColumnIndex: 0,
    endColumnIndex: 2,
  })
  expect(req.updateCells.fields).toBe("userEnteredFormat")
  expect(req.updateCells.rows[0].values).toEqual([
    { userEnteredFormat: { backgroundColor: { red: 1, green: 0, blue: 0 } } },
    {},
  ])
})

test("Tab append USER_ENTERED passes valueInputOption through", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.append({ values: ["=A1+B1"], inputFormat: "USER_ENTERED" })

  expect(mockAppend.mock.calls[0]?.[0].valueInputOption).toBe("USER_ENTERED")
})

test("Tab append with getRow returns Row at echoed index", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  // Default mockAppend echoes "Sheet1!A3:B3" → row 3.
  mockGetSpreadsheet.mockImplementation(() =>
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
                      { effectiveValue: { stringValue: "x" } },
                      { effectiveValue: { stringValue: "y" } },
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

  const result = await tab.append({ values: ["x", "y"], getRow: true })

  expect(result).not.toBeNull()
  expect(result).toHaveLength(2)
  expect(result?.[0].value).toBe("x")
  expect(result?.[1].value).toBe("y")
  expect(result?.getRowIndex()).toBe(3)
})

test("Tab append getRow reads the echoed row with a targeted range", async () => {
  const { sheet, mockAppend, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  // Echo a large index to prove the read targets the echoed row, not row 1.
  mockAppend.mockImplementation(() =>
    Promise.resolve({
      data: { updates: { updatedRange: "Sheet1!A6:A6" } },
    }),
  )
  mockGetSpreadsheet.mockImplementation(() =>
    Promise.resolve({
      data: {
        sheets: [
          {
            properties: { title: "Sheet1", sheetId: 0 },
            data: [{ rowData: [{ values: [] }] }],
          },
        ],
      },
    }),
  )

  const result = await tab.append({ values: ["x"], getRow: true })

  // Only one get — the targeted single-row read (no gridProperties pre-fetch).
  expect(mockGetSpreadsheet).toHaveBeenCalledTimes(1)
  expect(mockGetSpreadsheet.mock.calls[0]?.[0].ranges?.[0]).toBe(
    "Sheet1!A6:ZZZ6",
  )
  expect(result?.getRowIndex()).toBe(6)
})

test("Tab append getRow returns actual applied format from API response", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  mockGetSpreadsheet.mockImplementation(() =>
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
                      {
                        effectiveValue: { stringValue: "x" },
                        userEnteredFormat: {
                          backgroundColor: { red: 1, green: 0, blue: 0 },
                        },
                      },
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

  const result = await tab.append({
    values: [new Cell({ value: "x" })],
    getRow: true,
  })

  expect(result?.[0].format?.backgroundColor).toEqual({
    red: 1,
    green: 0,
    blue: 0,
  })
})

test("Tab append getRow=false does not read the appended row", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  await tab.append({ values: ["x"], getRow: false })

  expect(mockGetSpreadsheet).not.toHaveBeenCalled()
})

test("Tab append getRow returns API format, not input Cell format", async () => {
  const { sheet, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  mockGetSpreadsheet.mockImplementation(() =>
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
                      {
                        effectiveValue: { stringValue: "x" },
                        userEnteredFormat: {
                          backgroundColor: { red: 1, green: 0, blue: 0 },
                        },
                      },
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

  const inputCell = new Cell({
    value: "x",
    format: { backgroundColor: { red: 0, green: 0, blue: 1 } },
  })

  const result = await tab.append({ values: [inputCell], getRow: true })

  // Returned format comes from API (red), not the input Cell (blue)
  expect(result?.[0].format?.backgroundColor).toEqual({
    red: 1,
    green: 0,
    blue: 0,
  })
})

test("Tab append with trailing empty grid rows lands at the data row, not grid size", async () => {
  // Regression: previously the appended index came from gridProperties.rowCount
  // (grid size, e.g. 1000), not the last data row. values.append echoes the
  // real landing range, so a near-empty 1000-row grid still resolves correctly.
  const { sheet, mockAppend, mockGetSpreadsheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  // 5 data rows in a 1000-row grid → append lands at row 6.
  mockAppend.mockImplementation(() =>
    Promise.resolve({
      data: { updates: { updatedRange: "Sheet1!A6:A6" } },
    }),
  )
  mockGetSpreadsheet.mockImplementation(() =>
    Promise.resolve({
      data: {
        sheets: [
          {
            properties: {
              title: "Sheet1",
              sheetId: 0,
              gridProperties: { rowCount: 1000 },
            },
            data: [{ rowData: [{ values: [] }] }],
          },
        ],
      },
    }),
  )

  const result = await tab.append({ values: ["x"], getRow: true })

  expect(result?.getRowIndex()).toBe(6)
  expect(mockGetSpreadsheet.mock.calls[0]?.[0].ranges?.[0]).toBe(
    "Sheet1!A6:ZZZ6",
  )
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
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Email", "Age"])

  await tab.append({
    values: { Name: "John", Email: "john@test.com", Age: "30" },
  })

  expect(mockAppend.mock.calls[0]?.[0].requestBody.values).toEqual([
    ["John", "john@test.com", "30"],
  ])
})

test("Tab append with object validates against schema", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Email", "Age"])
  tab.setSchema(
    z.object({ Name: z.string(), Email: z.string(), Age: z.string() }),
  )

  await tab.append({
    values: { Name: "John", Email: "john@test.com", Age: "30" },
  })

  expect(mockAppend).toHaveBeenCalledTimes(1)
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
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  const cell1 = new Cell({ value: "x" })

  await tab.append({ values: [cell1, "y"] })

  expect(mockAppend.mock.calls[0]?.[0].requestBody.values).toEqual([["x", "y"]])
})

test("Tab append array validates against schema by column", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["name", "score"])
  tab.setSchema(z.object({ name: z.string(), score: z.number() }))

  // "abc" fails z.number() for score column (cannot coerce)
  expect(tab.append({ values: ["A", "abc"] })).rejects.toThrow(ValidationError)
})

test("Tab append array passes valid values", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["name", "score"])
  tab.setSchema(z.object({ name: z.string(), score: z.string() }))

  await tab.append({ values: ["Alice", "95"] })

  expect(mockAppend).toHaveBeenCalledTimes(1)
})

// --- header/schema mapping (header-mapping fix) ---

test("Tab setSchema preserves real sheet header order", () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["name", "age", "email"])

  tab.setSchema(
    z.object({ email: z.string(), name: z.string(), age: z.string() }),
  )

  expect(tab.getHeaders()).toEqual(["name", "age", "email"])
})

test("Tab _setHeaders throws when schema key missing from headers", () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab.setSchema(z.object({ A: z.string(), B: z.string(), C: z.string() }))

  expect(() => tab._setHeaders(["A", "B"])).toThrow(
    "Schema keys not found in sheet headers: C. Sheet headers: A, B",
  )
})

test("Tab setSchema throws when schema key missing from cached headers", () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Age"])

  expect(() =>
    tab.setSchema(z.object({ Name: z.string(), Missing: z.string() })),
  ).toThrow(/Schema keys not found in sheet headers: Missing/)
})

test("Tab append with object maps to header order, not schema key order", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["name", "age", "email"])
  tab.setSchema(
    z.object({ email: z.string(), name: z.string(), age: z.string() }),
  )

  await tab.append({
    values: { name: "John", age: "30", email: "john@test.com" },
  })

  expect(mockAppend.mock.calls[0]?.[0].requestBody.values).toEqual([
    ["John", "30", "john@test.com"],
  ])
})

test("Tab append with object and no schema preserves behavior", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["A", "B"])

  await tab.append({ values: { A: "x", B: "y" } })

  expect(mockAppend.mock.calls[0]?.[0].requestBody.values).toEqual([["x", "y"]])
})

test("Tab setSchema allows schema as a subset of headers", () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Score", "Extra"])

  expect(() =>
    tab.setSchema(z.object({ Name: z.string(), Score: z.string() })),
  ).not.toThrow()
  expect(tab.getHeaders()).toEqual(["Name", "Score", "Extra"])
})

test("Tab setSchema is case sensitive when matching headers", () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["Name", "Age"])

  expect(() =>
    tab.setSchema(z.object({ name: z.string(), age: z.string() })),
  ).toThrow(
    "Schema keys not found in sheet headers: name, age. Sheet headers: Name, Age",
  )
})

// --- schema coercion (schema-coerce fix) ---

test("Tab append with typed object passes number schema", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["age"])
  tab.setSchema(z.object({ age: z.number() }))

  await tab.append({ values: { age: 25 } })

  expect(mockAppend).toHaveBeenCalledTimes(1)
})

test("Tab append with invalid object value fails number schema", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["age"])
  tab.setSchema(z.object({ age: z.number() }))

  expect(tab.append({ values: { age: "invalid" } })).rejects.toThrow(
    ValidationError,
  )
})

test("Tab append array coerces numeric string to number schema", async () => {
  const { sheet, mockAppend } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["age"])
  tab.setSchema(z.object({ age: z.number() }))

  await tab.append({ values: ["42"] })
  await tab.append({ values: [25] })

  expect(mockAppend).toHaveBeenCalledTimes(2)
})

test("Tab append array rejects non-coercible string for number schema", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})
  tab._setHeaders(["age"])
  tab.setSchema(z.object({ age: z.number() }))

  expect(tab.append({ values: ["abc"] })).rejects.toThrow(ValidationError)
})

// --- delCell invalid end (Bug 5) ---

test("Tab delCell throws on invalid end address", async () => {
  const { sheet } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  expect(tab.delCell({ start: "A1", end: "not-valid" })).rejects.toThrow(
    /Invalid cell address: not-valid/,
  )
})

test("Tab delCell with valid start but invalid end does not silently use start", async () => {
  const { sheet, mockBatchUpdate } = makeMockSheet()
  const tab = new Tab(sheet, "Sheet1", 0, {})

  expect(tab.delCell({ start: "B2", end: "###" })).rejects.toThrow(
    /Invalid cell address/,
  )
  expect(mockBatchUpdate).not.toHaveBeenCalled()
})
