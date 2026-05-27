import {
  type AccountInfo,
  PublicClientApplication,
} from "@azure/msal-browser";

const clientId = import.meta.env.VITE_AZURE_CLIENT_ID;
const tenantId = import.meta.env.VITE_AZURE_TENANT_ID ?? "common";
const configuredRedirectUri = import.meta.env.VITE_AZURE_REDIRECT_URI;

function resolveRedirectUri(): string {
  if (!configuredRedirectUri) {
    return window.location.origin;
  }

  try {
    const configuredUrl = new URL(configuredRedirectUri);
    // When the dev server falls back to a different localhost port,
    // force redirect URI to current origin to keep MSAL cache coherent.
    if (configuredUrl.origin !== window.location.origin) {
      return window.location.origin;
    }

    return configuredRedirectUri;
  } catch {
    return window.location.origin;
  }
}

const redirectUri = resolveRedirectUri();

if (!clientId) {
  // This throws early in dev if auth config is missing.
  // eslint-disable-next-line no-console
  console.warn("VITE_AZURE_CLIENT_ID is not set. Sign-in will fail until configured.");
}

const msal = new PublicClientApplication({
  auth: {
    clientId: clientId ?? "",
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
});

const armScopes = ["https://management.azure.com/user_impersonation"];

let initialized = false;

export function getMsalConfigError(): string | null {
  if (!clientId) {
    return "Missing VITE_AZURE_CLIENT_ID. Create frontend/.env from frontend/.env.example and set your Entra App (SPA) Client ID.";
  }

  return null;
}

export async function initializeMsal(): Promise<void> {
  const configError = getMsalConfigError();
  if (configError) {
    throw new Error(configError);
  }

  if (initialized) {
    return;
  }

  await msal.initialize();
  let redirectResult = null;
  try {
    redirectResult = await msal.handleRedirectPromise();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error ?? "");
    // In popup-first flows, stale redirect cache artifacts should not block app startup.
    if (!message.includes("no_token_request_cache_error")) {
      throw error;
    }
  }

  if (redirectResult?.account) {
    msal.setActiveAccount(redirectResult.account);
  }

  const existing = msal.getActiveAccount() ?? msal.getAllAccounts()[0];
  if (existing) {
    msal.setActiveAccount(existing);
  }

  initialized = true;
}

export function getActiveAccount(): AccountInfo | null {
  return msal.getActiveAccount();
}

export async function signIn(): Promise<void> {
  const configError = getMsalConfigError();
  if (configError) {
    throw new Error(configError);
  }

  await initializeMsal();

  await msal.loginRedirect({
    scopes: armScopes,
    prompt: "select_account",
    redirectUri,
  });
}

export async function signOut(): Promise<void> {
  const account = msal.getActiveAccount();
  if (!account) {
    return;
  }

  await msal.logoutPopup({
    account,
  });
}

export async function acquireArmToken(): Promise<string> {
  const configError = getMsalConfigError();
  if (configError) {
    throw new Error(configError);
  }

  await initializeMsal();

  const account = msal.getActiveAccount() ?? msal.getAllAccounts()[0];
  if (!account) {
    throw new Error("Sign in to acquire an ARM token.");
  }

  msal.setActiveAccount(account);

  const result = await msal.acquireTokenSilent({
    account,
    scopes: armScopes,
  });

  return result.accessToken;
}
