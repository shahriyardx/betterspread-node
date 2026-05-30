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
