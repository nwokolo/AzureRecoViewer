import { useEffect, useMemo, useState } from "react";
import axios from "axios";
import {
  CartesianGrid,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import "./App.css";
import {
  acquireArmToken,
  getActiveAccount,
  getMsalConfigError,
  initializeMsal,
  signIn,
  signOut,
} from "./auth/msal";

type LookBackPeriod = "Last7Days" | "Last30Days" | "Last60Days";
type Term = "P1Y" | "P3Y";
type BenefitScope = "Single" | "Shared";
type SavingsPlanType = "compute" | "database" | "";
type SingleScopeTarget = "subscription" | "resourceGroup" | "";

type SimpleItem = {
  id?: string;
  name?: string;
  displayName?: string;
};

type RecommendationDetail = {
  commitmentAmount?: number;
  savingsAmount?: number;
  savingsPercentage?: number;
  averageUtilizationPercentage?: number;
};

type ChartPoint = {
  timestamp: string;
  hourIndex: number;
  cost: number;
  commitment?: number;
};

type Recommendation = {
  id?: string;
  armSkuName?: string;
  kind?: string;
  scope?: string;
  term?: string;
  lookBackPeriod?: string;
  currencyCode?: string;
  recommendationDetails?: RecommendationDetail;
  allRecommendationDetails?: RecommendationDetail[];
  chartSeries: ChartPoint[];
};

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL ?? "http://localhost:4000";

function App() {
  const [signedInUser, setSignedInUser] = useState("");
  const [armToken, setArmToken] = useState("");
  const [authLoading, setAuthLoading] = useState(false);

  const [billingAccounts, setBillingAccounts] = useState<SimpleItem[]>([]);
  const [billingProfiles, setBillingProfiles] = useState<SimpleItem[]>([]);
  const [selectedBillingAccountId, setSelectedBillingAccountId] = useState("");
  const [selectedBillingProfileId, setSelectedBillingProfileId] = useState("");
  const [savingsPlanType, setSavingsPlanType] = useState<SavingsPlanType>("");
  const [singleScopeTarget, setSingleScopeTarget] = useState<SingleScopeTarget>("");
  const [subscriptions, setSubscriptions] = useState<SimpleItem[]>([]);
  const [resourceGroups, setResourceGroups] = useState<SimpleItem[]>([]);
  const [selectedSingleSubscriptionId, setSelectedSingleSubscriptionId] = useState("");
  const [selectedSingleResourceGroup, setSelectedSingleResourceGroup] = useState("");

  const [lookBackPeriod, setLookBackPeriod] = useState<LookBackPeriod>("Last30Days");
  const [term, setTerm] = useState<Term>("P3Y");
  const [benefitScope, setBenefitScope] = useState<BenefitScope>("Shared");

  const [selectedRecommendation, setSelectedRecommendation] = useState<Recommendation | null>(null);
  const [loadingRecommendations, setLoadingRecommendations] = useState(false);
  const [loadingProfiles, setLoadingProfiles] = useState(false);
  const [hasEvaluatedFilters, setHasEvaluatedFilters] = useState(false);
  const [error, setError] = useState("");
  const msalConfigError = getMsalConfigError();

  const api = useMemo(
    () =>
      axios.create({
        baseURL: API_BASE_URL,
        headers: {
          Authorization: armToken ? `Bearer ${armToken}` : "",
        },
      }),
    [armToken],
  );

  const canPickAccount = Boolean(signedInUser && armToken);
  const canPickProfile = Boolean(canPickAccount && selectedBillingAccountId);
  const canPickPlanType = Boolean(canPickProfile && selectedBillingProfileId);
  const canPickSingleScopeTarget = Boolean(canPickPlanType && savingsPlanType && benefitScope === "Single");
  const hasValidSingleScopeSelection =
    benefitScope === "Shared"
      ? true
      : singleScopeTarget === "subscription"
        ? Boolean(selectedSingleSubscriptionId)
        : Boolean(selectedSingleSubscriptionId && selectedSingleResourceGroup);
  const canPickRecommendationFilters = Boolean(canPickPlanType && savingsPlanType && hasValidSingleScopeSelection);

  function getFriendlyAuthError(authError: unknown): string {
    const fallback = "Sign-in failed.";
    if (!(authError instanceof Error)) {
      return fallback;
    }

    if (authError.message.includes("no_token_request_cache_error")) {
      return "MSAL token cache mismatch. Ensure your Entra app has this exact localhost redirect URI, then refresh and sign in again.";
    }

    if (authError.message.includes("block_nested_popups")) {
      return "Nested popup sign-in was blocked. Close any auth popup windows and retry sign-in from the main app window.";
    }

    return authError.message || fallback;
  }

  async function refreshArmToken() {
    setAuthLoading(true);
    try {
      const token = await acquireArmToken();
      const account = getActiveAccount();
      setArmToken(token);
      setSignedInUser(account?.username ?? "");
      setError("");
    } catch (authError) {
      setError(getFriendlyAuthError(authError));
    } finally {
      setAuthLoading(false);
    }
  }

  async function handleSignIn() {
    setAuthLoading(true);
    try {
      await signIn();
    } catch (authError) {
      setError(getFriendlyAuthError(authError));
    } finally {
      setAuthLoading(false);
    }
  }

  async function handleSignOut() {
    await signOut();
    setArmToken("");
    setSignedInUser("");
    setBillingAccounts([]);
    setBillingProfiles([]);
    setSelectedBillingAccountId("");
    setSelectedBillingProfileId("");
    setSavingsPlanType("");
    setSingleScopeTarget("");
    setSubscriptions([]);
    setResourceGroups([]);
    setSelectedSingleSubscriptionId("");
    setSelectedSingleResourceGroup("");
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  async function loadBillingAccounts() {
    if (!armToken) {
      setError("Sign in first to load billing scopes.");
      return;
    }

    setError("");
    const response = await api.get<{ value: SimpleItem[] }>("/api/scopes/billing-accounts");
    setBillingAccounts(response.data.value ?? []);
  }

  async function loadBillingProfiles(billingAccountId: string) {
    if (!billingAccountId) {
      setBillingProfiles([]);
      return;
    }

    setLoadingProfiles(true);

    const response = await api.get<{ value: SimpleItem[] }>(
      `/api/scopes/billing-accounts/${encodeURIComponent(billingAccountId)}/billing-profiles`,
    );

    setBillingProfiles(response.data.value ?? []);
    setLoadingProfiles(false);
  }

  async function loadSubscriptions() {
    if (!armToken) {
      return;
    }

    const response = await api.get<{ value: SimpleItem[] }>("/api/scopes/subscriptions");
    setSubscriptions(response.data.value ?? []);
  }

  async function loadResourceGroups(subscriptionId: string) {
    if (!armToken || !subscriptionId) {
      setResourceGroups([]);
      return;
    }

    const response = await api.get<{ value: SimpleItem[] }>(
      `/api/scopes/subscriptions/${encodeURIComponent(subscriptionId)}/resource-groups`,
    );
    setResourceGroups(response.data.value ?? []);
  }

  function matchesPlanType(recommendation: Recommendation): boolean {
    const sku = (recommendation.armSkuName ?? "").toLowerCase();
    if (savingsPlanType === "compute") {
      return sku.includes("compute");
    }

    if (savingsPlanType === "database") {
      return (
        sku.includes("database") ||
        sku.includes("sql") ||
        sku.includes("postgres") ||
        sku.includes("mysql") ||
        sku.includes("mariadb")
      );
    }

    return false;
  }

  async function refreshRecommendations() {
    if (!armToken || !selectedBillingAccountId || !selectedBillingProfileId || !savingsPlanType) {
      setSelectedRecommendation(null);
      setHasEvaluatedFilters(false);
      return;
    }

    if (benefitScope === "Single" && !hasValidSingleScopeSelection) {
      setSelectedRecommendation(null);
      setHasEvaluatedFilters(false);
      return;
    }

    try {
      setLoadingRecommendations(true);
      setError("");

      const params: Record<string, string> = {
        lookBackPeriod,
        term,
        benefitScope,
      };

      if (benefitScope === "Shared") {
        params.scopeType = "billingProfile";
        params.billingAccountId = selectedBillingAccountId;
        params.billingProfileId = selectedBillingProfileId;
      } else if (singleScopeTarget === "subscription") {
        params.scopeType = "subscription";
        params.subscriptionId = selectedSingleSubscriptionId;
      } else {
        params.scopeType = "resourceGroup";
        params.subscriptionId = selectedSingleSubscriptionId;
        params.resourceGroup = selectedSingleResourceGroup;
      }

      const response = await api.get<{ value: Recommendation[] }>("/api/recommendations/benefit", {
        params,
      });

      const list = (response.data.value ?? []).filter(matchesPlanType);
      const ranked = [...list].sort((a, b) => {
        const savingsA = Number(a.recommendationDetails?.savingsAmount ?? 0);
        const savingsB = Number(b.recommendationDetails?.savingsAmount ?? 0);
        return savingsB - savingsA;
      });

      setSelectedRecommendation(ranked[0] ?? null);
      setHasEvaluatedFilters(true);
    } catch (requestError) {
      if (axios.isAxiosError(requestError)) {
        setError(requestError.response?.data?.error ?? requestError.message);
      } else {
        setError("Failed to load recommendations.");
      }
      setSelectedRecommendation(null);
      setHasEvaluatedFilters(true);
    } finally {
      setLoadingRecommendations(false);
    }
  }

  function resetDownstreamFromAccount(): void {
    setSelectedBillingProfileId("");
    setSavingsPlanType("");
    setBenefitScope("Shared");
    setLookBackPeriod("Last30Days");
    setTerm("P3Y");
    setSingleScopeTarget("");
    setSelectedSingleSubscriptionId("");
    setSelectedSingleResourceGroup("");
    setResourceGroups([]);
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  function resetDownstreamFromProfile(): void {
    setSavingsPlanType("");
    setBenefitScope("Shared");
    setLookBackPeriod("Last30Days");
    setTerm("P3Y");
    setSingleScopeTarget("");
    setSelectedSingleSubscriptionId("");
    setSelectedSingleResourceGroup("");
    setResourceGroups([]);
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  function resetDownstreamFromPlanType(): void {
    setBenefitScope("Shared");
    setLookBackPeriod("Last30Days");
    setTerm("P3Y");
    setSingleScopeTarget("");
    setSelectedSingleSubscriptionId("");
    setSelectedSingleResourceGroup("");
    setResourceGroups([]);
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  function resetDownstreamFromBenefitScope(): void {
    setSingleScopeTarget("");
    setSelectedSingleSubscriptionId("");
    setSelectedSingleResourceGroup("");
    setResourceGroups([]);
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  function onBenefitScopeChange(nextScope: BenefitScope): void {
    setBenefitScope(nextScope);
    resetDownstreamFromBenefitScope();
  }

  function onSingleScopeTargetChange(nextTarget: SingleScopeTarget): void {
    setSingleScopeTarget(nextTarget);
    setSelectedSingleSubscriptionId("");
    setSelectedSingleResourceGroup("");
    setResourceGroups([]);
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  function onSingleSubscriptionChange(nextValue: string): void {
    setSelectedSingleSubscriptionId(nextValue);
    setSelectedSingleResourceGroup("");
    setResourceGroups([]);
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  function onSingleResourceGroupChange(nextValue: string): void {
    setSelectedSingleResourceGroup(nextValue);
    setHasEvaluatedFilters(false);
    setSelectedRecommendation(null);
  }

  function onBillingAccountChange(nextValue: string): void {
    setSelectedBillingAccountId(nextValue);
    resetDownstreamFromAccount();
  }

  function onBillingProfileChange(nextValue: string): void {
    setSelectedBillingProfileId(nextValue);
    resetDownstreamFromProfile();
  }

  function onSavingsPlanTypeChange(nextValue: SavingsPlanType): void {
    setSavingsPlanType(nextValue);
    resetDownstreamFromPlanType();
  }

  useEffect(() => {
    initializeMsal()
      .then(() => {
        const existing = getActiveAccount();
        if (existing) {
          setSignedInUser(existing.username);
          refreshArmToken().catch(() => {
            setError("Token refresh failed. Sign in again.");
          });
        }
      })
      .catch((authError) => {
        setError(getFriendlyAuthError(authError));
      });
  }, []);

  useEffect(() => {
    if (!armToken) {
      return;
    }

    loadBillingAccounts().catch(() => {
      setError("Failed to load billing accounts. Check token permissions.");
    });

    loadSubscriptions().catch(() => {
      setError("Failed to load subscriptions.");
    });
  }, [armToken]);

  useEffect(() => {
    if (!selectedBillingAccountId) {
      setBillingProfiles([]);
      return;
    }

    loadBillingProfiles(selectedBillingAccountId).catch(() => {
      setError("Failed to load billing profiles for selected account.");
      setLoadingProfiles(false);
    });
  }, [selectedBillingAccountId]);

  useEffect(() => {
    if (benefitScope !== "Single") {
      return;
    }

    if (singleScopeTarget !== "resourceGroup") {
      return;
    }

    if (!selectedSingleSubscriptionId) {
      setResourceGroups([]);
      return;
    }

    loadResourceGroups(selectedSingleSubscriptionId).catch(() => {
      setError("Failed to load resource groups for selected subscription.");
    });
  }, [armToken, benefitScope, singleScopeTarget, selectedSingleSubscriptionId]);

  useEffect(() => {
    if (!canPickRecommendationFilters) {
      return;
    }

    refreshRecommendations().catch(() => {
      setError("Failed to refresh recommendation.");
    });
  }, [
    armToken,
    selectedBillingAccountId,
    selectedBillingProfileId,
    savingsPlanType,
    benefitScope,
    singleScopeTarget,
    selectedSingleSubscriptionId,
    selectedSingleResourceGroup,
    lookBackPeriod,
    term,
    hasValidSingleScopeSelection,
    canPickRecommendationFilters,
  ]);

  const chartData = useMemo(() => {
    if (!selectedRecommendation) {
      return [];
    }

    const commitmentAmount = Number(selectedRecommendation.recommendationDetails?.commitmentAmount ?? 0);

    return selectedRecommendation.chartSeries.map((point) => {
      return {
        ...point,
        commitment: commitmentAmount,
      };
    });
  }, [selectedRecommendation]);

  const recommendedCommitment = Number(selectedRecommendation?.recommendationDetails?.commitmentAmount ?? 0);

  return (
    <div className="page">
      <header className="hero">
        <p className="eyebrow">Azure Cost Management Explorer</p>
        <h1>Commitment Recommendation Viewer</h1>
        <p>
          Select a billing scope, filter recommendations by term and lookback, then overlay the
          selected commitment on hourly usage cost.
        </p>
      </header>

      <section className="panel">
        <h2>Step 1: Sign In</h2>
        <div className="auth-row">
          <div>
            <strong>Signed in:</strong> {signedInUser || "Not signed in"}
          </div>
          <div className="auth-actions">
            {!signedInUser ? (
              <button type="button" onClick={handleSignIn} disabled={authLoading || Boolean(msalConfigError)}>
                {authLoading ? "Signing in..." : "Sign in with Microsoft Entra"}
              </button>
            ) : (
              <>
                <button type="button" onClick={refreshArmToken} disabled={authLoading}>
                  {authLoading ? "Refreshing..." : "Refresh Token"}
                </button>
                <button type="button" onClick={handleSignOut} disabled={authLoading}>
                  Sign out
                </button>
              </>
            )}
          </div>
        </div>
        {msalConfigError ? <p className="error">{msalConfigError}</p> : null}
      </section>

      <section className="panel grid-3">
        <h2>Step 2: Billing Account and Profile</h2>
        <label>
          Billing Account
          <select
            value={selectedBillingAccountId}
            onChange={(event) => onBillingAccountChange(event.target.value)}
            disabled={!canPickAccount}
          >
            <option value="">Select an account</option>
            {billingAccounts.map((account) => (
              <option key={account.id} value={account.name ?? account.id}>
                {account.displayName ?? account.name ?? account.id}
              </option>
            ))}
          </select>
        </label>
        <label>
          Billing Profile
          <select
            value={selectedBillingProfileId}
            onChange={(event) => onBillingProfileChange(event.target.value)}
            disabled={!canPickProfile || loadingProfiles}
          >
            <option value="">{loadingProfiles ? "Loading profiles..." : "Select a profile"}</option>
            {billingProfiles.map((profile) => (
              <option key={profile.id} value={profile.name ?? profile.id}>
                {profile.displayName ?? profile.name ?? profile.id}
              </option>
            ))}
          </select>
        </label>
      </section>

      <section className="panel grid-3">
        <h2>Step 3: Savings Plan Type</h2>
        <label>
          Savings plan type
          <select
            value={savingsPlanType}
            onChange={(event) => onSavingsPlanTypeChange(event.target.value as SavingsPlanType)}
            disabled={!canPickPlanType}
          >
            <option value="">Select plan type</option>
            <option value="compute">Compute</option>
            <option value="database">Database</option>
          </select>
        </label>
      </section>

      <section className="panel grid-3">
        <h2>Step 4: Scope, Lookback, and Term</h2>
        <label>
          Benefit Scope
          <select
            value={benefitScope}
            onChange={(event) => onBenefitScopeChange(event.target.value as BenefitScope)}
            disabled={!canPickRecommendationFilters}
          >
            <option value="Shared">Shared</option>
            <option value="Single">Single</option>
          </select>
        </label>

        {benefitScope === "Single" ? (
          <label>
            Single scope target
            <select
              value={singleScopeTarget}
              onChange={(event) => onSingleScopeTargetChange(event.target.value as SingleScopeTarget)}
              disabled={!canPickSingleScopeTarget}
            >
              <option value="">Select single scope</option>
              <option value="subscription">Subscription</option>
              <option value="resourceGroup">Resource Group</option>
            </select>
          </label>
        ) : null}

        {benefitScope === "Single" && singleScopeTarget ? (
          <label>
            Subscription
            <select
              value={selectedSingleSubscriptionId}
              onChange={(event) => onSingleSubscriptionChange(event.target.value)}
              disabled={!canPickSingleScopeTarget}
            >
              <option value="">Select subscription</option>
              {subscriptions.map((subscription) => (
                <option key={subscription.id} value={subscription.id}>
                  {subscription.displayName ?? subscription.id}
                </option>
              ))}
            </select>
          </label>
        ) : null}

        {benefitScope === "Single" && singleScopeTarget === "resourceGroup" ? (
          <label>
            Resource Group
            <select
              value={selectedSingleResourceGroup}
              onChange={(event) => onSingleResourceGroupChange(event.target.value)}
              disabled={!selectedSingleSubscriptionId}
            >
              <option value="">Select resource group</option>
              {resourceGroups.map((group) => (
                <option key={group.id} value={group.name ?? group.id}>
                  {group.name ?? group.id}
                </option>
              ))}
            </select>
          </label>
        ) : null}

        <label>
          Lookback
          <select
            value={lookBackPeriod}
            onChange={(event) => setLookBackPeriod(event.target.value as LookBackPeriod)}
            disabled={!canPickRecommendationFilters}
          >
            <option value="Last7Days">Last 7 Days</option>
            <option value="Last30Days">Last 30 Days</option>
            <option value="Last60Days">Last 60 Days</option>
          </select>
        </label>

        <label>
          Term
          <select
            value={term}
            onChange={(event) => setTerm(event.target.value as Term)}
            disabled={!canPickRecommendationFilters}
          >
            <option value="P1Y">1 Year</option>
            <option value="P3Y">3 Years</option>
          </select>
        </label>
      </section>

      {error ? <p className="error">{error}</p> : null}

      {hasEvaluatedFilters && !loadingRecommendations && !selectedRecommendation ? (
        <section className="panel">
          <p className="error">
            No recommendation available for the selected scope and term combination.
          </p>
        </section>
      ) : null}

      {selectedRecommendation ? (
        <>
          <section className="panel list-panel">
            <div className="section-head">
              <h2>Recommendation</h2>
              <span>{loadingRecommendations ? "Refreshing..." : "1 available"}</span>
            </div>
            <div className="recommendation active">
              <strong>{selectedRecommendation.armSkuName ?? "Unknown SKU"}</strong>
              <span>{selectedRecommendation.kind ?? "Unknown kind"}</span>
              <span>
                Savings: {selectedRecommendation.recommendationDetails?.savingsAmount?.toFixed(2) ?? "0.00"}{" "}
                {selectedRecommendation.currencyCode ?? ""}
              </span>
              <span>
                Recommended Commitment/hr: {selectedRecommendation.recommendationDetails?.commitmentAmount?.toFixed(4) ?? "0.0000"}
              </span>
            </div>
          </section>

          <section className="panel chart-panel">
            <h2>Current Hourly Spend vs Recommended Hourly Commitment</h2>
            <div className="metrics">
              <article>
                <h3>Recommended Commitment/hr</h3>
                <p>{recommendedCommitment.toFixed(4)}</p>
              </article>
              <article>
                <h3>Savings %</h3>
                <p>{selectedRecommendation.recommendationDetails?.savingsPercentage?.toFixed(2) ?? "0.00"}%</p>
              </article>
              <article>
                <h3>Utilization %</h3>
                <p>{selectedRecommendation.recommendationDetails?.averageUtilizationPercentage?.toFixed(2) ?? "0.00"}%</p>
              </article>
            </div>
            <div className="chart-wrap">
              <ResponsiveContainer width="100%" height={360}>
                <LineChart data={chartData}>
                  <CartesianGrid strokeDasharray="4 4" stroke="#2f3c53" />
                  <XAxis
                    dataKey="hourIndex"
                    tickFormatter={(value) => `${value + 1}`}
                    stroke="#7f95bd"
                  />
                  <YAxis stroke="#7f95bd" />
                  <Tooltip
                    formatter={(value, name) => [
                      Number(value ?? 0).toFixed(4),
                      name === "On-demand spend" ? "On-demand spend" : "Recommended Commitment",
                    ]}
                    labelFormatter={(value) => `Hour ${Number(value) + 1}`}
                  />
                  <Line
                    type="monotone"
                    dataKey="cost"
                    stroke="#0ec5a2"
                    strokeWidth={2.5}
                    dot={false}
                    name="On-demand spend"
                  />
                  <Line
                    type="monotone"
                    dataKey="commitment"
                    stroke="#f4a259"
                    strokeWidth={2.5}
                    strokeDasharray="8 5"
                    dot={false}
                    name="Recommended Commitment"
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </section>
        </>
      ) : null}
    </div>
  );
}

export default App;
