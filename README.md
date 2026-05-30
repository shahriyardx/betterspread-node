# betterspread

TypeScript Google Sheets wrapper with typed rows, cell formatting, and Zod schema validation.

## Install

```bash
bun add betterspread
```

## Quick Start

```ts
import { Connection, Sheet } from "betterspread"

const conn = new Connection({ credentialsPath: "./credentials.json" })
const sheet = new Sheet("My Spreadsheet", conn)
await sheet.open()
const tab = await sheet.getTab("Sheet1")

// Read all values
const rows = await tab.values()
console.log(rows[0][0].value) // first cell of first row

// Read with format data
const rowsWithFormat = await tab.values()
rowsWithFormat[0][0].format // Format instance or null

// Read a specific row
const row = await tab.getRow({ serialNo: 1 })

// Read a specific cell
const cell = await tab.getCell({ cellName: "B3" })

// Append rows
await tab.append({ values: ["Alice", "alice@test.com", "30"] })
await tab.append({ values: { Name: "Bob", Email: "bob@test.com" } })

// Delete rows
await tab.delRow({ start: 3 })
await tab.delRow({ start: 3, end: 5 })

// Delete cells with shift
await tab.delCell({ start: "B2" })
await tab.delCell({ start: "A1", end: "C3", shift: "left" })
```

## Rows

```ts
const row = await tab.getRow({ serialNo: 1 })

// Update cells
await row.update({ values: ["x", "y", "z"] })
await row.update({ values: { Name: "Alice", Age: "25" } })
await row.update({ values: ["x", "y"], inputFormat: "USER_ENTERED" })

// Style entire row
await row.style(new Format({ bgColor: "#ff0000" }))

// Append cell to end of row
await row.appendCell("new-value")
await row.appendCell(["x", "y", "z"])

// Clear row contents
await row.clear()

// Refresh from API
await row.refetch()

// Delete the row
await row.delete()

// Access cells by column label
row.get("A") // Cell
row.get("B") // Cell
```

## Cells

```ts
const cell = await tab.getCell({ cellName: "B3" })

// Read
console.log(cell.value) // "cell value"
console.log(cell.header) // column header or ""

// Update
await cell.update({ value: "new-value" })
await cell.update({ value: "new-value", format: new Format({ bold: true }) })
await cell.update({ value: "=SUM(A1:A10)", inputFormat: "USER_ENTERED" })

// Style
await cell.style(new Format({ bgColor: "#00ff00", italic: true }))

// Clear
await cell.clear()

// Delete (shift left or up)
await cell.delete("left")
await cell.delete("up")
```

## Zod Schema Validation

```ts
import { z } from "zod"

const UserSchema = z.object({
  Name: z.string(),
  Email: z.string().email(),
  Age: z.number(),
})

tab.setSchema(UserSchema)

// Validates on append
await tab.append({ values: { Name: "Alice", Email: "alice@test.com", Age: 30 } })
// Throws ValidationError
await tab.append({ values: { Name: "Bob", Email: "bob", Age: "not-a-number" } })

// Also validates array append and row updates
await tab.append({ values: ["Alice", "alice@test.com", "30"] }) // Age is string, schema expects number
```

## Cell Formatting

```ts
import { Format } from "betterspread"

const fmt = new Format({
  bgColor: "#ff0000",
  bold: true,
  italic: true,
  fontSize: 14,
  horizontalAlign: "center",
  borders: {
    top: { style: "solid", color: { red: 0, green: 0, blue: 0 } },
    bottom: { style: "dashed" },
  },
})

await cell.style(fmt)

// Extend existing format
const base = new Format({ bgColor: "#ff0000", bold: true })
const extended = base.extend({ bgColor: "#00ff00", italic: true })
```

## Formula Support

```ts
// Write formulas
await cell.update({ value: "=SUM(A1:A10)", inputFormat: "USER_ENTERED" })
await tab.append({ values: ["=B1*C1"], inputFormat: "USER_ENTERED" })

// Read formulas (returns formula string, not computed value)
await tab.values({ valueRenderOption: "FORMULA" })
```

## Value Render Options

```ts
// FORMULA: returns formula strings instead of computed values
await tab.values({ valueRenderOption: "FORMULA" })
await tab.getRow({ serialNo: 1, valueRenderOption: "FORMULA" })
await tab.getCell({ cellName: "A1", valueRenderOption: "FORMULA" })
```

## License

MIT
