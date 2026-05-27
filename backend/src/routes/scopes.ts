import { Router } from "express";
import { armListAll, armUrl } from "../services/armClient";
import { getBearerToken } from "../middleware/requireArmToken";

type BillingAccount = {
  id?: string;
  name?: string;
  properties?: {
    displayName?: string;
  };
};

type BillingProfile = {
  id?: string;
  name?: string;
  properties?: {
    displayName?: string;
  };
};

type Subscription = {
  subscriptionId?: string;
  displayName?: string;
};

type ResourceGroup = {
  id?: string;
  name?: string;
};

export const scopesRouter = Router();

scopesRouter.get("/billing-accounts", async (req, res, next) => {
  try {
    const token = getBearerToken(req);
    const accounts = await armListAll<BillingAccount>(
      armUrl("/providers/Microsoft.Billing/billingAccounts?api-version=2024-04-01"),
      token,
    );

    res.json({
      value: accounts.map((account) => ({
        id: account.id,
        name: account.name,
        displayName: account.properties?.displayName ?? account.name,
      })),
    });
  } catch (error) {
    next(error);
  }
});

scopesRouter.get("/billing-accounts/:billingAccountId/billing-profiles", async (req, res, next) => {
  try {
    const token = getBearerToken(req);
    const billingAccountId = encodeURIComponent(req.params.billingAccountId);
    const path = `/providers/Microsoft.Billing/billingAccounts/${billingAccountId}/billingProfiles?api-version=2024-04-01`;

    const profiles = await armListAll<BillingProfile>(armUrl(path), token);

    res.json({
      value: profiles.map((profile) => ({
        id: profile.id,
        name: profile.name,
        displayName: profile.properties?.displayName ?? profile.name,
      })),
    });
  } catch (error) {
    next(error);
  }
});

scopesRouter.get("/subscriptions", async (req, res, next) => {
  try {
    const token = getBearerToken(req);
    const subscriptions = await armListAll<Subscription>(
      armUrl("/subscriptions?api-version=2022-12-01"),
      token,
    );

    res.json({
      value: subscriptions.map((subscription) => ({
        id: subscription.subscriptionId,
        displayName: subscription.displayName,
      })),
    });
  } catch (error) {
    next(error);
  }
});

scopesRouter.get("/subscriptions/:subscriptionId/resource-groups", async (req, res, next) => {
  try {
    const token = getBearerToken(req);
    const subscriptionId = encodeURIComponent(req.params.subscriptionId);
    const path = `/subscriptions/${subscriptionId}/resourcegroups?api-version=2022-09-01`;
    const resourceGroups = await armListAll<ResourceGroup>(armUrl(path), token);

    res.json({
      value: resourceGroups.map((resourceGroup) => ({
        id: resourceGroup.id,
        name: resourceGroup.name,
      })),
    });
  } catch (error) {
    next(error);
  }
});
