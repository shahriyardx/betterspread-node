# Row

Extends `Array<Cell>`. Contains cells with metadata and parent references.

```ts
import { Row } from "betterspread"
```

## Methods

### update()

```ts
update(opts: RowUpdateOpts): Promise<void>
```

Updates cells in this row via batch update.

#### RowUpdateOpts

| Property | Type | Description |
|----------|------|-------------|
| `values` | `unknown[] \| Record<string, unknown>` | Array or object keyed by header |
| `inputFormat?` | `ValueInputOption` | `"RAW"` (default) or `"USER_ENTERED"` |

### clear()

```ts
clear(): Promise<void>
```

Clears all cell values in this row.

### style()

```ts
style(style: Format): Promise<void>
```

Applies a Format to the entire row.

### appendCell()

```ts
appendCell(value: unknown | unknown[]): Promise<void>
```

Appends cell(s) to the end of this row.

### refetch()

```ts
refetch(): Promise<void>
```

Re-reads row data from the API and replaces local cells.

### delete()

```ts
delete(): Promise<void>
```

Deletes the entire row from the sheet.

### get()

```ts
get(col: string): Cell
```

Returns cell by column label. Returns virtual cell if column doesn't exist.

### Accessors

```ts
row.getTab(): TabInstance
row.getRowIndex(): number
```

## Properties

| Index access | Returns |
|-------------|---------|
| `row[0]` | First Cell |
| `row["A"]` | Cell at column A (via Proxy) |
| `row.length` | Number of cells |
