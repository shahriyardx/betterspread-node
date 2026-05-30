# Formula Support

## Writing Formulas

Use `inputFormat: "USER_ENTERED"` — Sheets API parses values starting with `=` as formulas:

```ts
// Cell update
await cell.update({
  value: "=SUM(A1:A10)",
  inputFormat: "USER_ENTERED",
})

// Row update
await row.update({
  values: ["=B1*C1", "=A1+B1"],
  inputFormat: "USER_ENTERED",
})

// Append
await tab.append({
  values: ["=SUM(D2:D10)", "=AVERAGE(E2:E10)"],
  inputFormat: "USER_ENTERED",
})
```

Plain values (no `=` prefix) are stored as strings even with `USER_ENTERED`.

## Reading Formulas

Use `valueRenderOption: "FORMULA"` to get formula strings instead of computed values:

```ts
// Returns formula strings
const rows = await tab.values({ valueRenderOption: "FORMULA" })

// Single row
const row = await tab.getRow({
  serialNo: 1,
  valueRenderOption: "FORMULA",
})

// Single cell
const cell = await tab.getCell({
  cellName: "A1",
  valueRenderOption: "FORMULA",
})
```

Default behavior (no render option) returns computed values via `FORMATTED_VALUE`.
