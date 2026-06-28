# Zod Schema Validation

Attach a Zod schema to a Tab to validate data on append and update.

## Setup

```ts
import { z } from "zod"
import { ValidationError } from "betterspread"

const UserSchema = z.object({
  Name: z.string(),
  Email: z.string().email(),
  Age: z.number(),
})

tab.setSchema(UserSchema)
```

Schema keys must exist in the tab's cached headers (case-sensitive). `setSchema`
throws if a key has no matching header. The schema does **not** reorder columns —
values map to columns by header name, preserving real sheet column order.

## Type Coercion

Cells store strings, so array and cell paths coerce before validating: a numeric
string passes a `z.number()` column, `"true"`/`"false"` pass `z.boolean()`, and a
date string passes `z.date()`. Non-coercible values still throw.

```ts
// "30" coerces to 30 → passes Age: z.number()
await tab.append({ values: ["Alice", "alice@test.com", "30"] })

// "abc" cannot coerce → throws ValidationError
await tab.append({ values: ["Alice", "alice@test.com", "abc"] })
```

## Validation on Append

```ts
// Object append — validates full record against schema
await tab.append({
  values: { Name: "Alice", Email: "alice@test.com", Age: 30 },
})

// Throws ValidationError
await tab.append({
  values: { Name: "Bob", Email: "bob", Age: "not-a-number" },
})
```

## Validation on Row Update

```ts
await row.update({ values: { Name: "Alice", Email: "alice@test.com", Age: 30 } })

// Also validates array updates by column position
await row.update({ values: ["Alice", "alice@test.com", "30"] })
```

## Validation on Cell Update

```ts
const cell = await tab.getCell({ cellName: "B2" })
// If header maps to schema field, value is validated
await cell.update({ value: "abc" })
```

## Catching Errors

```ts
try {
  await tab.append({ values: { Name: "X", Age: "bad" } })
} catch (err) {
  if (err instanceof ValidationError) {
    console.log(err.message) // human-readable
    console.log(err.cause) // original ZodError
  }
}
```
