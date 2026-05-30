import type { sheets_v4 } from "@googleapis/sheets"
import type { drive_v3 } from "@googleapis/drive"
import type { Connection } from "./connection"
import { Tab } from "./tab"
import { SheetNotFoundError, TabNotFoundError } from "./types"

export class Sheet {
  private sheetName: string
  private connection: Connection
  private folderId?: string
  private _id: string | null = null
  private _client: sheets_v4.Sheets | null = null

  constructor(sheetName: string, connection: Connection, folderId?: string) {
    this.sheetName = sheetName
    this.connection = connection
    this.folderId = folderId
  }

  getClient(): sheets_v4.Sheets {
    if (!this._client) throw new Error("Sheet not opened. Call open() first.")
    return this._client
  }

  getId(): string {
    if (!this._id) throw new Error("Sheet not opened. Call open() first.")
    return this._id
  }

  async open(): Promise<void> {
    const client = await this.connection.getSheetsClient()
    this._client = client

    const id = await this.resolveSheetId(client)
    if (!id) throw new SheetNotFoundError(this.sheetName)

    this._id = id
  }

  private async resolveSheetId(
    client: sheets_v4.Sheets,
  ): Promise<string | null> {
    if (this.folderId) {
      return this.resolveByName()
    }

    const directId = await this.tryDirectId(client)
    if (directId) return directId

    return this.resolveByName()
  }

  private async tryDirectId(client: sheets_v4.Sheets): Promise<string | null> {
    try {
      const res = await client.spreadsheets.get({
        spreadsheetId: this.sheetName,
      })
      return res.data.spreadsheetId ?? null
    } catch {
      return null
    }
  }

  private async resolveByName(): Promise<string | null> {
    let driveClient: drive_v3.Drive
    try {
      driveClient = await this.connection.getDriveClient()
    } catch {
      throw new Error(
        "Drive API not available. Pass spreadsheet ID instead of name, or enable Drive API for your service account.",
      )
    }

    const queryParts: string[] = [
      `name = '${this.sheetName.replace(/'/g, "\\'")}'`,
      "mimeType = 'application/vnd.google-apps.spreadsheet'",
    ]

    if (this.folderId) {
      queryParts.push(`'${this.folderId}' in parents`)
    }

    try {
      const res = await driveClient.files.list({
        q: queryParts.join(" and "),
        fields: "files(id, name)",
        pageSize: 1,
      })

      const files = res.data.files ?? []
      if (files.length === 0) return null

      return files[0]?.id ?? null
    } catch {
      throw new Error(
        "Drive API request failed. Pass spreadsheet ID instead of name, or enable Drive API for your service account.",
      )
    }
  }

  async getTab(tabName: string): Promise<Tab> {
    await this.ensureOpen()
    const sheets = await this.fetchSheets()
    const ws = sheets.find((s) => s.properties?.title === tabName)

    if (!ws) throw new TabNotFoundError(tabName)

    return new Tab(
      this,
      ws.properties?.title ?? tabName,
      ws.properties?.sheetId ?? 0,
      ws.properties?.gridProperties ?? {},
    )
  }

  async tabs(opts?: { excludeHidden?: boolean }): Promise<Tab[]> {
    await this.ensureOpen()
    const sheets = await this.fetchSheets()

    let filtered = sheets
    if (opts?.excludeHidden) {
      filtered = sheets.filter((s) => !s.properties?.hidden)
    }

    return filtered.map((ws) => {
      const props = ws.properties
      return new Tab(
        this,
        props?.title ?? `Sheet-${props?.sheetId ?? 0}`,
        props?.sheetId ?? 0,
        props?.gridProperties ?? {},
      )
    })
  }

  private async ensureOpen(): Promise<void> {
    if (!this._id) await this.open()
  }

  private async fetchSheets(): Promise<sheets_v4.Schema$Sheet[]> {
    if (!this._client || !this._id) {
      throw new Error("Sheet not opened. Call open() first.")
    }
    const res = await this._client.spreadsheets.get({
      spreadsheetId: this._id,
    })
    return res.data.sheets ?? []
  }
}
