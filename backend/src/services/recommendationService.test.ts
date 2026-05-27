import { describe, expect, it } from "vitest";
import {
  buildBillingScope,
  buildFilters,
  recommendationQuerySchema,
  toHourlyPoints,
} from "./recommendationService";

describe("recommendationService", () => {
  it("builds billing profile scope", () => {
    const query = recommendationQuerySchema.parse({
      scopeType: "billingProfile",
      billingAccountId: "1234",
      billingProfileId: "abcd",
      benefitScope: "Shared",
      lookBackPeriod: "Last30Days",
      term: "P1Y",
    });

    expect(buildBillingScope(query)).toBe(
      "/providers/Microsoft.Billing/billingAccounts/1234/billingProfiles/abcd",
    );
  });

  it("throws for missing required scope fields", () => {
    const query = recommendationQuerySchema.parse({
      scopeType: "resourceGroup",
      subscriptionId: "sub-1",
      benefitScope: "Shared",
      lookBackPeriod: "Last30Days",
      term: "P1Y",
    });

    expect(() => buildBillingScope(query)).toThrow(
      "subscriptionId and resourceGroup are required for resourceGroup scope.",
    );
  });

  it("composes ARM filter string", () => {
    const query = recommendationQuerySchema.parse({
      scopeType: "billingAccount",
      billingAccountId: "acct-1",
      benefitScope: "Single",
      lookBackPeriod: "Last7Days",
      term: "P3Y",
    });

    expect(buildFilters(query)).toBe(
      "properties/scope eq 'Single' AND properties/lookBackPeriod eq 'Last7Days' AND properties/term eq 'P3Y'",
    );
  });

  it("maps charges into hourly points with timestamps", () => {
    const points = toHourlyPoints("2026-01-01T00:00:00Z", [1, 2, 3]);

    expect(points).toHaveLength(3);
    expect(points[0]).toEqual({
      timestamp: "2026-01-01T00:00:00.000Z",
      hourIndex: 0,
      cost: 1,
    });
    expect(points[2]).toEqual({
      timestamp: "2026-01-01T02:00:00.000Z",
      hourIndex: 2,
      cost: 3,
    });
  });
});
