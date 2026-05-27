import { Router } from "express";
import { getBearerToken } from "../middleware/requireArmToken";
import {
  fetchBenefitRecommendations,
  recommendationQuerySchema,
  toHourlyPoints,
} from "../services/recommendationService";

export const recommendationsRouter = Router();

recommendationsRouter.get("/benefit", async (req, res, next) => {
  try {
    const parsedQuery = recommendationQuerySchema.parse(req.query);
    const token = getBearerToken(req);

    const filteredRecommendations = await fetchBenefitRecommendations(parsedQuery, token);

    const recommendations = filteredRecommendations.map((resource) => {
      const properties = resource.properties ?? {};
      const usageCharges = properties.usage?.charges ?? [];
      const recommendationDetails = properties.recommendationDetails ?? {};
      const allRecommendationDetails = properties.allRecommendationDetails?.value ?? [];
      const chartSeries = toHourlyPoints(properties.firstConsumptionDate, usageCharges);
      const commitmentAmount = Number(recommendationDetails.commitmentAmount ?? 0);
      const overlaySeries = chartSeries.map((point) => ({
        ...point,
        commitment: commitmentAmount,
      }));

      return {
        id: resource.id,
        name: resource.name,
        kind: resource.kind,
        armSkuName: properties.armSkuName,
        currencyCode: properties.currencyCode,
        scope: properties.scope,
        lookBackPeriod: properties.lookBackPeriod,
        term: properties.term,
        firstConsumptionDate: properties.firstConsumptionDate,
        lastConsumptionDate: properties.lastConsumptionDate,
        totalHours: properties.totalHours ?? chartSeries.length,
        costWithoutBenefit: properties.costWithoutBenefit,
        recommendationDetails,
        allRecommendationDetails,
        chartSeries,
        overlaySeries,
      };
    });

    res.json({
      value: recommendations,
      summary: {
        totalRecommendations: recommendations.length,
        availableSkus: Array.from(
          new Set(
            recommendations
              .map((recommendation) => recommendation.armSkuName)
              .filter((sku): sku is string => Boolean(sku)),
          ),
        ).sort(),
      },
    });
  } catch (error) {
    next(error);
  }
});
