# Cell Formatting

## Creating Formats

```ts
import { Format } from "betterspread"

const fmt = new Format({
  bgColor: "#ff0000",
  bold: true,
  italic: true,
  fontSize: 14,
  fontFamily: "Arial",
  horizontalAlign: "center",
  verticalAlign: "middle",
})
```

## Borders

```ts
const fmt = new Format({
  borders: {
    top: { style: "solid", color: { red: 0, green: 0, blue: 0 } },
    bottom: { style: "dashed" },
    left: { style: "dotted" },
    right: { style: "double" },
  },
})
```

Border styles: `solid`, `dashed`, `dotted`, `double`, `none`, `medium`, `mediumDashed`, `mediumDotted`, `mediumSolid`, `thick`.

## Applying Styles

```ts
// To a cell
await cell.style(new Format({ bgColor: "#00ff00" }))

// To a row
await row.style(new Format({ bgColor: "#ff0000", italic: true }))
```

## Extending Formats

`extend()` creates a new Format with merged properties — original unchanged:

```ts
const base = new Format({ bgColor: "#ff0000", bold: true, fontSize: 14 })
const extended = base.extend({ bgColor: "#00ff00", italic: true })
// extended: bold=true, italic=true, fontSize=14, bgColor="#00ff00"
// base unchanged
```

## All Format Options

| Shorthand | Maps to |
|-----------|---------|
| `bgColor` | `backgroundColor` (hex → RGB) |
| `textColor` | `textFormat.foregroundColor` |
| `bold` | `textFormat.bold` |
| `italic` | `textFormat.italic` |
| `strikethrough` | `textFormat.strikethrough` |
| `underline` | `textFormat.underline` |
| `fontFamily` | `textFormat.fontFamily` |
| `fontSize` | `textFormat.fontSize` |
| `link` | `textFormat.link` |
| `horizontalAlign` | `horizontalAlignment` |
| `verticalAlign` | `verticalAlignment` |

Full CellFormat properties also accepted directly: `backgroundColor`, `backgroundColorStyle`, `horizontalAlignment`, `verticalAlignment`, `textFormat`, `numberFormat`, `padding`, `textDirection`, `textRotation`, `wrapStrategy`, `hyperlinkDisplayType`.
