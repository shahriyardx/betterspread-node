import { test, expect } from "bun:test"
import { Connection } from "../connection"
import { ValueError } from "../types"

test("Connection throws ValueError with no credentials", () => {
  expect(() => new Connection({})).toThrow(ValueError)
})

test("Connection throws ValueError with both credentials", () => {
  expect(
    () =>
      new Connection({
        credentialsPath: "./a.json",
        credentialsDict: { client_email: "a", private_key: "b" },
      }),
  ).toThrow(ValueError)
})

test("Connection throws ValueError message says exactly one", () => {
  expect(() => new Connection({})).toThrow(
    /exactly one of credentialsPath or credentialsDict/i,
  )
})

test("Connection accepts credentialsPath", () => {
  const c = new Connection({ credentialsPath: "./creds.json" })
  expect(c).toBeDefined()
})

test("Connection accepts credentialsDict", () => {
  const c = new Connection({
    credentialsDict: { client_email: "a", private_key: "b" },
  })
  expect(c).toBeDefined()
})
