# Connection

Manages Google Cloud service account authentication.

```ts
import { Connection } from "betterspread"
```

## Constructor

```ts
new Connection(opts: ConnectionOptions)
```

### ConnectionOptions

| Property | Type | Description |
|----------|------|-------------|
| `credentialsPath` | `string` | Path to service account JSON file |
| `credentialsDict` | `Record<string, unknown>` | Inline service account credentials |

Exactly one of `credentialsPath` or `credentialsDict` is required.

## Methods

### getSheetsClient()

```ts
getSheetsClient(): Promise<sheets_v4.Sheets>
```

Returns authenticated Google Sheets API client.

### getDriveClient()

```ts
getDriveClient(): Promise<drive_v3.Drive>
```

Returns authenticated Google Drive API client (used for resolving spreadsheet names).
