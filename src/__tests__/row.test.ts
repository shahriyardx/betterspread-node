import { test, expect, mock } from "bun:test"
import { z } from "zod"
import { Row } from "../row"
import { Cell } from "../cell"
import { ValidationError } from "../types"
import { Format } from "../format"
import type { sheets_v4 } from "@googleapis/sheets"

function makeMockTab() {
  const mockUpdate = mock(() => Promise.resolve({ data: {} }))
  const mockClear = mock(() => Promise.resolve({ data: {} }))
  const mockBatchUpdate = mock(() => Promise.resolve({ data: {} }))
  const mockGet = mock(() =>
    Promise.resolve({ data: { values: [["x", "y"]] } }),
  )

  const mockClient = {
    spreadsheets: {
      values: {
        update: mockUpdate,
        clear: mockClear,
        get: mockGet,
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
  }

  return { tab, mockClient, mockUpdate, mockClear, mockBatchUpdate, mockGet }
}

function makeCells(tab: any, values: string[], rowIndex: number): Cell[] {
  return values.map((v, i) => {
    const label = String.fromCharCode(65 + i)
    return new Cell({
      value: v,
      label,
      rowIndex,
      cellIndex: i,
      tab,
      row: null,
    })
  })
}

test("Row extends Array and contains Cells", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b", "c"], 1)
  const row = new Row(cells, tab, 1)

  expect(row).toHaveLength(3)
  expect(row[0]).toBeInstanceOf(Cell)
  expect(row[0].value).toBe("a")
  expect(row[1].value).toBe("b")
})

test("Row cells have back-reference to parent Row", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  expect(row[0].row).toBe(row)
  expect(row[1].row).toBe(row)
})

test("Row update calls API with correct range", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  await row.update({ values: ["x", "y"] })

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.updateCells.range.startRowIndex).toBe(0)
  expect(req.updateCells.range.startColumnIndex).toBe(0)
  expect(req.updateCells.range.endColumnIndex).toBe(2)
  expect(req.updateCells.rows[0].values).toEqual([
    { userEnteredValue: { stringValue: "x" } },
    { userEnteredValue: { stringValue: "y" } },
  ])
})

test("Row clear calls API", async () => {
  const { tab, mockClear } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  await row.clear()
  expect(mockClear).toHaveBeenCalledTimes(1)
})

test("Row style calls batchUpdate", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 2)
  const row = new Row(cells, tab, 2)

  await row.style(
    new Format({ backgroundColor: { red: 0.5, green: 0.5, blue: 0.5 } }),
  )

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.repeatCell.range.startRowIndex).toBe(1)
  expect(req.repeatCell.range.endRowIndex).toBe(2)
})

test("Row appendCell appends single value", async () => {
  const { tab, mockUpdate } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  await row.appendCell("c")

  expect(mockUpdate.mock.calls[0][0].range).toContain("C1")
  expect(mockUpdate.mock.calls[0][0].requestBody.values).toEqual([["c"]])
})

test("Row appendCell appends array of values", async () => {
  const { tab, mockUpdate } = makeMockTab()
  const cells = makeCells(tab, ["a"], 1)
  const row = new Row(cells, tab, 1)

  await row.appendCell(["y", "z"])

  expect(mockUpdate.mock.calls[0][0].requestBody.values).toEqual([["y", "z"]])
})

test("Row refetch replaces cell contents", async () => {
  const { tab, mockGet } = makeMockTab()
  const cells = makeCells(tab, ["old1", "old2"], 1)
  const row = new Row(cells, tab, 1)

  mockGet.mockImplementation(() =>
    Promise.resolve({ data: { values: [["new1", "new2", "new3"]] } }),
  )

  await row.refetch()

  expect(row).toHaveLength(3)
  expect(row[0].value).toBe("new1")
  expect(row[2].value).toBe("new3")
})

test("Row delete calls batchUpdate with deleteDimension", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cells = makeCells(tab, ["a"], 3)
  const row = new Row(cells, tab, 3)

  await row.delete()

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.deleteDimension.range.dimension).toBe("ROWS")
  expect(req.deleteDimension.range.startIndex).toBe(2)
  expect(req.deleteDimension.range.endIndex).toBe(3)
})

test("Row custom inspect shows index and cell count", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b", "c"], 2)
  const row = new Row(cells, tab, 2)

  const inspected = row[Symbol.for("nodejs.util.inspect.custom")]()
  expect(inspected).toBe("Row(index=2, cells=3)")
})

test("Row get by column label returns correct cell", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b", "c"], 2)
  const row = new Row(cells, tab, 2)

  expect(row.get("A")?.value).toBe("a")
  expect(row.get("B")?.value).toBe("b")
  expect(row.get("C")?.value).toBe("c")
})

test("Row get by lowercase column label works", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["x", "y"], 1)
  const row = new Row(cells, tab, 1)

  expect(row.get("a")?.value).toBe("x")
})

test("Row get returns virtual cell for out of range column", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  const cell = row.get("Z")
  expect(cell.value).toBe("")
  expect(cell.label).toBe("Z")
  expect(cell.rowIndex).toBe(1)
  expect(cell.cellIndex).toBe(25)
})

test("Row numeric index returns virtual cell for out of bounds", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  const cell = row[5]
  expect(cell?.value).toBe("")
  expect(cell?.label).toBe("F")
  expect(cell?.rowIndex).toBe(1)
  expect(cell?.cellIndex).toBe(5)
})

test("Row numeric index handles negative", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  expect(row[-1]).toBeUndefined()
})

test("Row update with Cell objects applies format", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 2)
  const row = new Row(cells, tab, 2)

  const styledCell = new Cell({
    value: "x",
    format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
  })

  await row.update({ values: [styledCell, "y"] })

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.updateCells.rows[0].values).toEqual([
    {
      userEnteredValue: { stringValue: "x" },
      userEnteredFormat: { backgroundColor: { red: 1, green: 0, blue: 0 } },
    },
    { userEnteredValue: { stringValue: "y" } },
  ])
})

test("Row update with object maps headers to columns", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const tabWithHeaders = {
    ...tab,
    getHeaders: () => ["Name", "Email", "Age"],
  }
  const cells = makeCells(tabWithHeaders, ["old", "", ""], 1)
  const row = new Row(cells, tabWithHeaders, 1)

  await row.update({ values: { Name: "Alice", Age: "25" } })

  const req = mockBatchUpdate.mock.calls[0][0].requestBody.requests[0]
  expect(req.updateCells.rows[0].values).toEqual([
    { userEnteredValue: { stringValue: "Alice" } },
    { userEnteredValue: { stringValue: "" } },
    { userEnteredValue: { stringValue: "25" } },
  ])
})

test("Row update with object validates against schema", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const tabWithSchema = {
    ...tab,
    getHeaders: () => ["Name", "Score"],
    getSchema: () => z.object({ Name: z.string(), Score: z.number() }),
  }
  const cells = makeCells(tabWithSchema, ["", ""], 1)
  const row = new Row(cells, tabWithSchema, 1)

  await expect(
    row.update({ values: { Name: "Bob", Score: "bad" } }),
  ).rejects.toThrow(ValidationError)

  expect(mockBatchUpdate).not.toHaveBeenCalled()
})

test("Row update array validates against schema by column", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const tabWithSchema = {
    ...tab,
    getHeaders: () => ["name", "score"],
    getSchema: () => z.object({ name: z.string(), score: z.number() }),
  }
  const cells = makeCells(tabWithSchema, ["", ""], 1)
  const row = new Row(cells, tabWithSchema, 1)

  // "abc" fails z.number() for score column (index 1)
  await expect(row.update({ values: ["Bob", "abc"] })).rejects.toThrow(
    ValidationError,
  )

  expect(mockBatchUpdate).not.toHaveBeenCalled()
})

test("Row update array passes valid values", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const tabWithSchema = {
    ...tab,
    getHeaders: () => ["name", "score"],
    getSchema: () => z.object({ name: z.string(), score: z.string() }),
  }
  const cells = makeCells(tabWithSchema, ["", ""], 1)
  const row = new Row(cells, tabWithSchema, 1)

  await row.update({ values: ["Alice", "95"] })

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})

test("Row.map() returns plain Array and calls callback for each cell", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b", "c"], 1)
  const row = new Row(cells, tab, 1)

  const result = row.map((cell) => cell.value.toUpperCase())

  expect(result).not.toBeInstanceOf(Row)
  expect(result).toBeInstanceOf(Array)
  expect(result).toEqual(["A", "B", "C"])
})

test("Row.filter() returns plain Array with filtered cells", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "", "c"], 1)
  const row = new Row(cells, tab, 1)

  const result = row.filter((cell) => cell.value !== "")

  expect(result).not.toBeInstanceOf(Row)
  expect(result).toBeInstanceOf(Array)
  expect(result).toHaveLength(2)
  expect(result[0]?.value).toBe("a")
})

test("Row.slice() returns plain Array of sliced cells", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b", "c"], 1)
  const row = new Row(cells, tab, 1)

  const result = row.slice(1, 3)

  expect(result).not.toBeInstanceOf(Row)
  expect(result).toBeInstanceOf(Array)
  expect(result).toHaveLength(2)
})

test("Row.splice() returns plain Array of deleted elements", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b", "c"], 1)
  const row = new Row(cells, tab, 1)

  const result = row.splice(1, 1)

  expect(result).not.toBeInstanceOf(Row)
  expect(result).toBeInstanceOf(Array)
  expect(result).toHaveLength(1)
  expect(result[0]?.value).toBe("b")
  expect(row).toHaveLength(2)
})

test("Row.concat() returns plain Array with concatenated cells", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)
  const newCells = makeCells(tab, ["c"], 1)

  const result = row.concat(newCells)

  expect(result).not.toBeInstanceOf(Row)
  expect(result).toBeInstanceOf(Array)
  expect(result).toHaveLength(3)
})

test("Row in-place mutations still work (push/sort)", async () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["c", "a"], 1)
  const row = new Row(cells, tab, 1)

  row.sort((a, b) => a.value.localeCompare(b.value))
  expect(row[0]?.value).toBe("a")
  expect(row[1]?.value).toBe("c")

  await row.appendCell("b")
  expect(row).toHaveLength(3)
  expect(row[2]?.value).toBe("b")
})

test("Row Proxy virtual cells still work with array methods", () => {
  const { tab } = makeMockTab()
  const cells = makeCells(tab, ["a", "b"], 1)
  const row = new Row(cells, tab, 1)

  const cell = row[5]
  expect(cell?.value).toBe("")
  expect(cell?.cellIndex).toBe(5)

  const result = row.map((c) => c.value)
  expect(result).toEqual(["a", "b"])
})

test("Row update array coerces typed values against schema", async () => {
  const { tab, mockBatchUpdate } = makeMockTab()
  const tabWithSchema = {
    ...tab,
    getHeaders: () => ["age", "flag"],
    getSchema: () => z.object({ age: z.number(), flag: z.boolean() }),
  }
  const cells = makeCells(tabWithSchema, ["", ""], 1)
  const row = new Row(cells, tabWithSchema, 1)

  // 30 is already a number; "true" coerces to boolean
  await row.update({ values: [30, "true"] })

  expect(mockBatchUpdate).toHaveBeenCalledTimes(1)
})
