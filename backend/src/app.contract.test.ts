import request from "supertest";
import { describe, expect, it } from "vitest";
import { createApp } from "./app";

describe("API contract", () => {
  it("returns health payload", async () => {
    const app = createApp();

    const response = await request(app).get("/api/health");

    expect(response.status).toBe(200);
    expect(response.body).toEqual({ status: "ok" });
  });

  it("requires bearer token for recommendations", async () => {
    const app = createApp();

    const response = await request(app).get("/api/recommendations/benefit");

    expect(response.status).toBe(401);
    expect(response.body.error).toContain("Missing Bearer token");
  });
});
