# Format

Cell formatting with shorthand properties mapped to Google Sheets CellFormat.

```ts
import { Format } from "betterspread"
```

## Constructor

```ts
new Format(opts?: FormatOptions)
```

All properties optional. See [Formatting guide](/guide/formatting) for full option list.

## Instance Methods

### toCellFormat()

```ts
toCellFormat(): CellFormat
```

Converts to Google Sheets API `CellFormat` object.

### extend()

```ts
extend(opts: FormatOptions): Format
```

Creates new Format with overrides. Original unchanged.

```ts
const base = new Format({ bgColor: "#ff0000", bold: true })
const extended = base.extend({ bgColor: "#00ff00" })
```

### fromCellFormat() (static)

```ts
Format.fromCellFormat(format: CellFormat): Format
```

Creates a Format from a Google Sheets API CellFormat response object.

## Shorthand Properties

| Property | Type | Maps to |
|----------|------|---------|
| `bgColor` | `string` | Hex color → `backgroundColor` |
| `textColor` | `string` | Hex color → `textFormat.foregroundColor` |
| `bold` | `boolean` | `textFormat.bold` |
| `italic` | `boolean` | `textFormat.italic` |
| `strikethrough` | `boolean` | `textFormat.strikethrough` |
| `underline` | `boolean` | `textFormat.underline` |
| `fontFamily` | `string` | `textFormat.fontFamily` |
| `fontSize` | `number` | `textFormat.fontSize` |
| `link` | `Link` | `textFormat.link` |
| `horizontalAlign` | `"left" \| "center" \| "right"` | `horizontalAlignment` |
| `verticalAlign` | `"top" \| "middle" \| "bottom"` | `verticalAlignment` |

Full CellFormat properties also available directly: `backgroundColor`, `backgroundColorStyle`, `borders`, `horizontalAlignment`, `verticalAlignment`, `hyperlinkDisplayType`, `numberFormat`, `padding`, `textDirection`, `textFormat`, `textRotation`, `wrapStrategy`. These take precedence over shorthands when both are set.
