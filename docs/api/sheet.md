# Sheet

Represents a Google Spreadsheet. Resolves by name or ID.

```ts
import { Sheet } from "betterspread"
```

## Constructor

```ts
new Sheet(sheetName: string, connection: Connection, folderId?: string)
```

| Param | Type | Description |
|-------|------|-------------|
| `sheetName` | `string` | Spreadsheet name or ID |
| `connection` | `Connection` | Authenticated connection |
| `folderId` | `string` | (optional) Restrict name search to this Drive folder |

## Methods

### open()

```ts
open(): Promise<void>
```

Resolves the spreadsheet ID and initializes the API client. Must be called before other methods.

### getTab()

```ts
getTab(tabName: string): Promise<Tab>
```

Gets a Tab by sheet name. Throws `TabNotFoundError` if not found.

### tabs()

```ts
tabs(opts?: { excludeHidden?: boolean }): Promise<Tab[]>
```

Returns all tabs. Optionally exclude hidden sheets.

### getClient()

```ts
getClient(): sheets_v4.Sheets
```

Returns the Sheets API client. Throws if `open()` not called.

### getId()

```ts
getId(): string
```

Returns the resolved spreadsheet ID. Throws if `open()` not called.
