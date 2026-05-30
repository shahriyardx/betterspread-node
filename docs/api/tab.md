# Tab

Represents a worksheet (sheet tab) within a spreadsheet.

```ts
import { Tab } from "betterspread"
```

## Methods

### values()

```ts
values(opts?: ValuesOpts): Promise<Row[]>
```

Reads rows from the tab. Default returns all rows with cell format data.

#### ValuesOpts

| Property | Type | Description |
|----------|------|-------------|
| `range` | `string` | A1 range (e.g. `"Sheet1!A1:B10"`) |
| `valueRenderOption` | `ValueRenderOption` | `"FORMATTED_VALUE"` (default), `"UNFORMATTED_VALUE"`, or `"FORMULA"` |

### getRow()

```ts
getRow(opts: GetRowOpts): Promise<Row>
```

Reads a single row by number.

#### GetRowOpts

| Property | Type | Description |
|----------|------|-------------|
| `serialNo` | `number` | Row number (1-indexed) |
| `valueRenderOption?` | `ValueRenderOption` | Render option for values |

### getCell()

```ts
getCell(opts: GetCellOpts): Promise<Cell>
```

Reads a single cell by A1 address.

#### GetCellOpts

| Property | Type | Description |
|----------|------|-------------|
| `cellName` | `string` | A1 notation (e.g. `"B3"`) |
| `valueRenderOption?` | `ValueRenderOption` | Render option (default `"FORMATTED_VALUE"`) |

### append()

```ts
append(opts: AppendOpts): Promise<Row | null>
```

Appends a row to the end of the tab.

#### AppendOpts

| Property | Type | Description |
|----------|------|-------------|
| `values` | `unknown[] \| Record<string, unknown>` | Array of values or object keyed by header |
| `inputFormat?` | `ValueInputOption` | `"RAW"` (default) or `"USER_ENTERED"` |
| `getRow?` | `boolean` | If true, returns the appended Row |

### delRow()

```ts
delRow(opts: DelRowOpts): Promise<void>
```

Deletes a row or range of rows.

#### DelRowOpts

| Property | Type | Description |
|----------|------|-------------|
| `start` | `number` | Start row (1-indexed) |
| `end?` | `number` | End row (exclusive) |

### delCell()

```ts
delCell(opts: DelCellOpts): Promise<void>
```

Deletes a cell or range with shift.

#### DelCellOpts

| Property | Type | Description |
|----------|------|-------------|
| `start` | `string` | Start cell (A1 notation) |
| `end?` | `string` | End cell (A1 notation) |
| `shift?` | `"up" \| "left"` | Shift direction (default `"up"`) |

### setSchema()

```ts
setSchema(schema: z.ZodObject): this
```

Attaches a Zod schema for validation. Returns `this` for chaining.

### getSchema()

```ts
getSchema(): z.ZodObject | null
```

### getHeaders()

```ts
getHeaders(): string[]
```

Returns cached column headers from last `values()` call.
