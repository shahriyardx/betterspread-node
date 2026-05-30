# Getting Started

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
console.log(rows[0][0].value)

// Append a row
await tab.append({ values: ["Alice", "alice@test.com", "30"] })

// Append with object (requires headers cached)
await tab.append({
  values: { Name: "Bob", Email: "bob@test.com", Age: "25" },
})
```

## Credentials

You need a Google Cloud service account with Sheets API enabled.

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a service account and download the JSON key
3. Share your spreadsheet with the service account email
4. Pass the JSON path to `Connection`
