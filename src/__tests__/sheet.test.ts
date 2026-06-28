import { test, expect, mock } from "bun:test"
import { Sheet } from "../sheet"

function makeMockConnection(opts?: {
  directFails?: boolean
  driveFails?: boolean
}) {
  const mockDriveFilesList = mock(() =>
    opts?.driveFails
      ? Promise.reject(new Error("Drive API not enabled"))
      : Promise.resolve({
          data: { files: [{ id: "abc123", name: "My Sheet" }] },
        }),
  )

  const mockSpreadsheetsGet = mock(() =>
    opts?.directFails
      ? Promise.reject(new Error("Not found"))
      : Promise.resolve({ data: { spreadsheetId: "abc123" } }),
  )

  const mockSpreadsheetsGetFull = mock(() =>
    Promise.resolve({
      data: {
        spreadsheetId: "abc123",
        sheets: [
          {
            properties: {
              title: "Sheet1",
              sheetId: 0,
              gridProperties: { rowCount: 100, columnCount: 26 },
            },
          },
          {
            properties: {
              title: "HiddenSheet",
              sheetId: 1,
              hidden: true,
              gridProperties: { rowCount: 50, columnCount: 10 },
            },
          },
        ],
      },
    }),
  )

  const mockSheetsClient = {
    spreadsheets: { get: mockSpreadsheetsGet },
  }

  const mockDriveClient = {
    files: { list: mockDriveFilesList },
  }

  const connection = {
    getSheetsClient: () => Promise.resolve(mockSheetsClient),
    getDriveClient: () =>
      opts?.driveFails
        ? Promise.reject(new Error("Drive API not enabled"))
        : Promise.resolve(mockDriveClient),
  }

  return {
    connection,
    mockDriveFilesList,
    mockSpreadsheetsGet,
    mockSpreadsheetsGetFull,
    mockDriveClient,
  }
}

// Direct ID path (no Drive needed)
test("Sheet open resolves spreadsheet ID directly", async () => {
  const { connection, mockDriveFilesList } = makeMockConnection()
  const sheet = new Sheet("abc123", connection as any)

  await sheet.open()

  expect(sheet.getId()).toBe("abc123")
  // Drive API never called
  expect(mockDriveFilesList).toHaveBeenCalledTimes(0)
})

test("Sheet open with direct ID skips Drive API entirely", async () => {
  const { connection, mockDriveFilesList } = makeMockConnection({
    driveFails: true,
  })
  const sheet = new Sheet("abc123", connection as any)

  // Should succeed even though Drive fails, because direct ID works
  await sheet.open()

  expect(sheet.getId()).toBe("abc123")
  expect(mockDriveFilesList).toHaveBeenCalledTimes(0)
})

// Name resolution path (Drive needed)
test("Sheet open falls back to Drive search when direct ID fails", async () => {
  const { connection, mockDriveFilesList } = makeMockConnection({
    directFails: true,
  })
  const sheet = new Sheet("My Sheet", connection as any)

  await sheet.open()

  expect(sheet.getId()).toBe("abc123")
  expect(mockDriveFilesList).toHaveBeenCalledTimes(1)
  const q = mockDriveFilesList.mock.calls[0]![0]!.q
  expect(q).toContain("name = 'My Sheet'")
  expect(q).toContain("mimeType = 'application/vnd.google-apps.spreadsheet'")
})

test("Sheet open throws when both direct and Drive fail", async () => {
  const { connection, mockDriveFilesList } = makeMockConnection({
    directFails: true,
  })
  mockDriveFilesList.mockImplementation(() =>
    Promise.resolve({ data: { files: [] } }),
  )
  const sheet = new Sheet("Missing", connection as any)

  expect(sheet.open()).rejects.toThrow(/not found/i)
})

test("Sheet open gives helpful error when Drive not available", async () => {
  const { connection } = makeMockConnection({
    directFails: true,
    driveFails: true,
  })
  const sheet = new Sheet("My Sheet", connection as any)

  expect(sheet.open()).rejects.toThrow(/enable Drive API/i)
})

// folderId path
test("Sheet open with folderId skips direct ID, uses Drive", async () => {
  const { connection, mockSpreadsheetsGet, mockDriveFilesList } =
    makeMockConnection()
  const sheet = new Sheet("My Sheet", connection as any, "folder123")

  await sheet.open()

  // Direct ID should NOT be attempted
  expect(mockSpreadsheetsGet).toHaveBeenCalledTimes(0)
  expect(mockDriveFilesList).toHaveBeenCalledTimes(1)
  const q = mockDriveFilesList.mock.calls[0]![0]!.q
  expect(q).toContain("'folder123' in parents")
})

test("Sheet open with folderId and no Drive access throws helpful error", async () => {
  const { connection } = makeMockConnection({ driveFails: true })
  const sheet = new Sheet("My Sheet", connection as any, "folder123")

  expect(sheet.open()).rejects.toThrow(/enable Drive API/i)
})

// Post-open operations
test("Sheet getTab returns Tab with correct name", async () => {
  const { connection, mockSpreadsheetsGet } = makeMockConnection()
  // Override to return full sheet data including tabs
  mockSpreadsheetsGet.mockImplementation(() =>
    Promise.resolve({
      data: {
        spreadsheetId: "abc123",
        sheets: [
          {
            properties: {
              title: "Sheet1",
              sheetId: 0,
              gridProperties: { rowCount: 100, columnCount: 26 },
            },
            hidden: false,
          },
        ],
      },
    }),
  )

  const sheet = new Sheet("abc123", connection as any)
  const tab = await sheet.getTab("Sheet1")

  expect(tab.getTitle()).toBe("Sheet1")
  expect(tab.getWorksheetId()).toBe(0)
})

test("Sheet getTab throws TabNotFoundError for missing tab", async () => {
  const { connection } = makeMockConnection()
  const sheet = new Sheet("abc123", connection as any)

  expect(sheet.getTab("NonExistent")).rejects.toThrow(/not found/i)
})

test("Sheet tabs returns all tabs", async () => {
  const { connection, mockSpreadsheetsGet } = makeMockConnection()
  mockSpreadsheetsGet.mockImplementation(() =>
    Promise.resolve({
      data: {
        spreadsheetId: "abc123",
        sheets: [
          { properties: { title: "Sheet1", sheetId: 0 } },
          { properties: { title: "HiddenSheet", sheetId: 1, hidden: true } },
        ],
      },
    }),
  )

  const sheet = new Sheet("abc123", connection as any)
  const tabs = await sheet.tabs()

  expect(tabs).toHaveLength(2)
  expect(tabs[0]!.getTitle()).toBe("Sheet1")
  expect(tabs[1]!.getTitle()).toBe("HiddenSheet")
})

test("Sheet tabs with excludeHidden returns visible only", async () => {
  const { connection, mockSpreadsheetsGet } = makeMockConnection()
  mockSpreadsheetsGet.mockImplementation(() =>
    Promise.resolve({
      data: {
        spreadsheetId: "abc123",
        sheets: [
          { properties: { title: "Sheet1", sheetId: 0 } },
          { properties: { title: "HiddenSheet", sheetId: 1, hidden: true } },
        ],
      },
    }),
  )

  const sheet = new Sheet("abc123", connection as any)
  const tabs = await sheet.tabs({ excludeHidden: true })

  expect(tabs).toHaveLength(1)
  expect(tabs[0]!.getTitle()).toBe("Sheet1")
})

test("Sheet open escapes single quotes in sheet name", async () => {
  const { connection, mockSpreadsheetsGet, mockDriveFilesList } =
    makeMockConnection({ directFails: true })
  const sheet = new Sheet("O'Brien Sheet", connection as any)

  await sheet.open()

  const q = mockDriveFilesList.mock.calls[0]![0]!.q
  expect(q).toContain("name = 'O\\'Brien Sheet'")
})

test("Sheet custom inspect before open shows name only", () => {
  const { connection } = makeMockConnection()
  const sheet = new Sheet("My Sheet", connection as any)

  const inspected = sheet[Symbol.for("nodejs.util.inspect.custom")]()
  expect(inspected).toBe('Sheet(name="My Sheet")')
})

test("Sheet custom inspect after open shows name and id", async () => {
  const { connection } = makeMockConnection()
  const sheet = new Sheet("abc123", connection as any)

  await sheet.open()

  const inspected = sheet[Symbol.for("nodejs.util.inspect.custom")]()
  expect(inspected).toBe('Sheet(name="abc123", id="abc123")')
})

test("Sheet open escapes backslash in sheet name", async () => {
  const { connection, mockDriveFilesList } = makeMockConnection({
    directFails: true,
  })
  // Actual name contains a single backslash: a\b
  const sheet = new Sheet("a\\b", connection as any)

  await sheet.open()

  const q = mockDriveFilesList.mock.calls[0]![0]!.q
  // Backslash must be escaped to two backslashes: name = 'a\\b'
  expect(q).toContain("name = 'a\\\\b'")
})

test("Sheet open escapes backslash before quote (no malformed sequence)", async () => {
  const { connection, mockDriveFilesList } = makeMockConnection({
    directFails: true,
  })
  // Actual name contains backslash followed by quote: a\'b
  const sheet = new Sheet("a\\'b", connection as any)

  await sheet.open()

  const q = mockDriveFilesList.mock.calls[0]![0]!.q
  // Backslash escaped first (\\), then quote (\'): name = 'a\\\'b'
  expect(q).toContain("name = 'a\\\\\\'b'")
})

test("Sheet open preserves mimeType part when escaping name", async () => {
  const { connection, mockDriveFilesList } = makeMockConnection({
    directFails: true,
  })
  const sheet = new Sheet("a\\b", connection as any)

  await sheet.open()

  const q = mockDriveFilesList.mock.calls[0]![0]!.q
  expect(q).toContain("mimeType = 'application/vnd.google-apps.spreadsheet'")
})
