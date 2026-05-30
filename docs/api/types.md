# Types

Shared interfaces and error classes.

## Type Aliases

```ts
type ValueInputOption = "RAW" | "USER_ENTERED"
type ValueRenderOption = "FORMATTED_VALUE" | "UNFORMATTED_VALUE" | "FORMULA"
```

## Interfaces

### TabInstance

```ts
interface TabInstance {
  getClient(): sheets_v4.Sheets
  getSheetId(): string
  getTitle(): string
  getWorksheetId(): number
  getHeaders(): string[]
  getSchema(): z.ZodObject | null
}
```

### RowInstance

```ts
interface RowInstance {
  getTab(): TabInstance
  getRowIndex(): number
}
```

### SheetInstance

```ts
interface SheetInstance {
  getClient(): sheets_v4.Sheets
  getId(): string
}
```

## Error Classes

| Class | Description |
|-------|-------------|
| `ValueError` | Invalid configuration |
| `ValidationError` | Zod schema validation failed. Has `.cause` (original ZodError) |
| `SheetNotFoundError` | Spreadsheet not found |
| `TabNotFoundError` | Worksheet tab not found |

```ts
import { ValidationError } from "betterspread"

try {
  await tab.append({ values: badData })
} catch (err) {
  if (err instanceof ValidationError) {
    console.error(err.message)
    console.error(err.cause) // ZodError
  }
}
```
