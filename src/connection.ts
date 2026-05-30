import { sheets } from "@googleapis/sheets"
import type { sheets_v4 } from "@googleapis/sheets"
import { drive } from "@googleapis/drive"
import type { drive_v3 } from "@googleapis/drive"
import { JWT } from "google-auth-library"
import { ValueError } from "./types"

export interface ConnectionOptions {
  credentialsPath?: string
  credentialsDict?: Record<string, unknown>
}

export class Connection {
  private _sheetsClient: sheets_v4.Sheets | null = null
  private _driveClient: drive_v3.Drive | null = null
  private _auth: JWT | null = null
  private opts: ConnectionOptions

  constructor(opts: ConnectionOptions) {
    if (!opts.credentialsPath && !opts.credentialsDict) {
      throw new ValueError(
        "Exactly one of credentialsPath or credentialsDict must be supplied",
      )
    }
    if (opts.credentialsPath && opts.credentialsDict) {
      throw new ValueError(
        "Exactly one of credentialsPath or credentialsDict must be supplied",
      )
    }
    this.opts = opts
  }

  private async getAuth(): Promise<JWT> {
    if (this._auth) return this._auth

    let creds: Record<string, unknown>
    if (this.opts.credentialsPath) {
      const file = Bun.file(this.opts.credentialsPath)
      creds = await file.json()
    } else {
      creds = this.opts.credentialsDict!
    }

    this._auth = new JWT({
      email: creds.client_email as string,
      key: creds.private_key as string,
      scopes: [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
      ],
    })

    return this._auth
  }

  async getSheetsClient(): Promise<sheets_v4.Sheets> {
    if (this._sheetsClient) return this._sheetsClient
    const auth = await this.getAuth()
    this._sheetsClient = sheets({ version: "v4", auth: auth as any })
    return this._sheetsClient
  }

  async getDriveClient(): Promise<drive_v3.Drive> {
    if (this._driveClient) return this._driveClient
    const auth = await this.getAuth()
    this._driveClient = drive({ version: "v3", auth: auth as any })
    return this._driveClient
  }
}
