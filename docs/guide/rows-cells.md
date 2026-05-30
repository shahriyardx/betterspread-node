# Rows & Cells

## Reading Rows

```ts
// All rows
const rows = await tab.values()

// Specific row by number
const row = await tab.getRow({ serialNo: 1 })

// With render option
const row = await tab.getRow({ serialNo: 1, valueRenderOption: "FORMULA" })
```

## Reading Cells

```ts
const cell = await tab.getCell({ cellName: "B3" })
console.log(cell.value) // cell contents
console.log(cell.header) // column header from cache
console.log(cell.label) // "B"
console.log(cell.rowIndex) // 3
```

## Updating Cells

```ts
await cell.update({ value: "new-value" })
await cell.update({
  value: "new-value",
  format: new Format({ bold: true }),
})
await cell.update({
  value: "=SUM(A1:A10)",
  inputFormat: "USER_ENTERED",
})
```

## Updating Rows

```ts
// Array updates all columns in order
await row.update({ values: ["x", "y", "z"] })

// Object updates named columns (requires cached headers)
await row.update({ values: { Name: "Alice", Age: "25" } })

// With formula input
await row.update({ values: ["=B1*C1"], inputFormat: "USER_ENTERED" })
```

## Virtual Cells

Accessing a column beyond row length returns a virtual empty cell:

```ts
const cell = row.get("Z") // virtual cell, value = ""
console.log(cell.value) // ""
console.log(cell.cellIndex) // 25
```

## Clearing & Deleting

```ts
await cell.clear()
await row.clear()
await row.delete()

await tab.delRow({ start: 3 })
await tab.delRow({ start: 3, end: 5 })
await tab.delCell({ start: "B2" })
await tab.delCell({ start: "A1", end: "C3", shift: "left" })
```
