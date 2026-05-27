import { z } from "zod";
import { armListAll, armUrl } from "./armClient";

export const recommendationQuerySchema = z.object({
  scopeType: z.enum(["billingAccount", "billingProfile", "subscription", "resourceGroup"]),
  billingAccountId: z.string().optional(),
  billingProfileId: z.string().optional(),
  subscriptionId: z.string().optional(),
  resourceGroup: z.string().optional(),
  lookBackPeriod: z.enum(["Last7Days", "Last30Days", "Last60Days"]).default("Last30Days"),
  term: z.enum(["P1Y", "P3Y"]).default("P1Y"),
  benefitScope: z.enum(["Single", "Shared"]).default("Shared"),
  sku: z.string().optional(),
});

export type RecommendationQuery = z.infer<typeof recommendationQuerySchema>;

export type RecommendationDetail = {
  commitmentAmount?: number;
  savingsAmount?: number;
  savingsPercentage?: number;
  averageUtilizationPercentage?: number;
  benefitCost?: number;
  overageCost?: number;
  wastageCost?: number;
  coveragePercentage?: number;
};

export type RecommendationUsage = {
  usageGrain?: string;
  charges?: number[];
};

export type RecommendationResource = {
  id?: string;
  name?: string;
  kind?: string;
  properties?: {
    armSkuName?: string;
    currencyCode?: string;
    lookBackPeriod?: string;
    term?: string;
    scope?: string;
    firstConsumptionDate?: string;
    lastConsumptionDate?: string;
    totalHours?: number;
    costWithoutBenefit?: number;
    recommendationDetails?: RecommendationDetail;
    allRecommendationDetails?: {
      value?: RecommendationDetail[];
    };
    usage?: RecommendationUsage;
  };
};

export type HourlyPoint = {
  timestamp: string;
  hourIndex: number;
  cost: number;
};

export function buildBillingScope(query: RecommendationQuery): string {
  switch (query.scopeType) {
    case "billingProfile": {
      if (!query.billingAccountId || !query.billingProfileId) {
        throw new Error("billingAccountId and billingProfileId are required for billingProfile scope.");
      }

      return `/providers/Microsoft.Billing/billingAccounts/${encodeURIComponent(query.billingAccountId)}/billingProfiles/${encodeURIComponent(query.billingProfileId)}`;
    }
    case "billingAccount": {
      if (!query.billingAccountId) {
        throw new Error("billingAccountId is required for billingAccount scope.");
      }

      return `/providers/Microsoft.Billing/billingAccounts/${encodeURIComponent(query.billingAccountId)}`;
    }
    case "subscription": {
      if (!query.subscriptionId) {
        throw new Error("subscriptionId is required for subscription scope.");
      }

      return `/subscriptions/${encodeURIComponent(query.subscriptionId)}`;
    }
    case "resourceGroup": {
      if (!query.subscriptionId || !query.resourceGroup) {
        throw new Error("subscriptionId and resourceGroup are required for resourceGroup scope.");
      }

      return `/subscriptions/${encodeURIComponent(query.subscriptionId)}/resourceGroups/${encodeURIComponent(query.resourceGroup)}`;
    }
    default:
      return "";
  }
}

export function buildFilters(query: RecommendationQuery): string {
  const filters = [
    `properties/scope eq '${query.benefitScope}'`,
    `properties/lookBackPeriod eq '${query.lookBackPeriod}'`,
    `properties/term eq '${query.term}'`,
  ];

  return filters.join(" AND ");
}

export function toHourlyPoints(firstConsumptionDate?: string, charges: number[] = []): HourlyPoint[] {
  const start = firstConsumptionDate ? new Date(firstConsumptionDate).getTime() : Number.NaN;

  return charges.map((charge, index) => {
    const timestamp = Number.isNaN(start)
      ? `hour-${index + 1}`
      : new Date(start + index * 60 * 60 * 1000).toISOString();

    return {
      timestamp,
      hourIndex: index,
      cost: Number(charge ?? 0),
    };
  });
}

export async function fetchBenefitRecommendations(
  query: RecommendationQuery,
  token: string,
): Promise<RecommendationResource[]> {
  const billingScope = buildBillingScope(query);
  const filter = buildFilters(query);
  const endpointPath = `${billingScope}/providers/Microsoft.CostManagement/benefitRecommendations`;

  const params = new URLSearchParams({
    "api-version": "2025-03-01",
    $filter: filter,
    $expand: "properties/usage,properties/allRecommendationDetails",
  });

  const allRecommendations = await armListAll<RecommendationResource>(
    armUrl(`${endpointPath}?${params.toString()}`),
    token,
  );

  if (!query.sku) {
    return allRecommendations;
  }

  const skuFilterLower = query.sku.toLowerCase();
  return allRecommendations.filter((resource) =>
    (resource.properties?.armSkuName ?? "").toLowerCase().includes(skuFilterLower),
  );
}
