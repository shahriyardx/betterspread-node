# Cell

Represents a single cell with value, metadata, and optional format.

```ts
import { Cell } from "betterspread"
```

## Constructor

```ts
new Cell(opts: CellOptions)
```

### CellOptions

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `value` | `string` | — | Cell value |
| `label?` | `string` | `"A"` | Column label |
| `rowIndex?` | `number` | `1` | Row number |
| `cellIndex?` | `number` | `0` | Column index |
| `tab?` | `TabInstance` | — | Parent Tab |
| `row?` | `RowInstance \| null` | `null` | Parent Row |
| `format?` | `CellFormat \| Format` | — | Cell formatting |

## Properties

| Property | Type | Description |
|----------|------|-------------|
| `value` | `string` (readonly) | Cell value |
| `label` | `string` (readonly) | Column label (e.g. "B") |
| `rowIndex` | `number` (readonly) | Row number |
| `cellIndex` | `number` (readonly) | Column index (0-based) |
| `tab` | `TabInstance \| undefined` (readonly) | Parent Tab |
| `row` | `RowInstance \| null` (readonly) | Parent Row |
| `format` | `Format \| null` | Cell format |
| `header` | `string` (getter) | Column header from Tab cache |

## Methods

### update()

```ts
update(opts: CellUpdateOpts): Promise<this>
```

Updates cell value and optional format. Returns `this` for chaining.

#### CellUpdateOpts

| Property | Type | Description |
|----------|------|-------------|
| `value` | `string` | New value |
| `format?` | `Format` | Optional cell format |
| `inputFormat?` | `ValueInputOption` | `"RAW"` (default) or `"USER_ENTERED"` |

### clear()

```ts
clear(): Promise<void>
```

Clears the cell value.

### style()

```ts
style(style: Format): Promise<void>
```

Applies a Format to this cell.

### delete()

```ts
delete(shift: "left" | "up" = "up"): Promise<void>
```

Deletes the cell, shifting remaining cells up or left. Default `"up"`, matching `Tab.delCell`.

### toString()

```ts
toString(): string
```

Returns cell value. Called implicitly on string coercion.

### toJSON()

```ts
toJSON(): string
```

Returns cell value for JSON serialization.
