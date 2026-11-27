import { InteractionRequiredAuthError, PublicClientApplication, type AccountInfo } from "@azure/msal-browser";

// Entra app registration "Mail AI". A public identifier, not a secret.
export const CLIENT_ID = "aff69668-2d41-411e-aa31-b677720065b3";

// Read-only mail access. Nothing else is requested.
const GRAPH_SCOPES = ["User.Read", "Mail.Read"];

const appUrl = new URL(import.meta.env.BASE_URL, window.location.origin).href; // .../ or .../app/

export const msal = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/common", // work/school + personal accounts
    redirectUri: new URL("redirect.html", appUrl).href,  // MSAL v5 redirect bridge page
    postLogoutRedirectUri: appUrl,
  },
  // Tokens stay in this browser only; localStorage keeps people signed in across tabs.
  cache: { cacheLocation: "localStorage" },
});

/** Initialise MSAL, finish a sign-in redirect if one is in progress, return the signed-in account. */
export async function initAuth(): Promise<AccountInfo | null> {
  await msal.initialize();
  const result = await msal.handleRedirectPromise();
  const account = result?.account ?? msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null;
  if (account) msal.setActiveAccount(account);
  return account;
}

export function signIn() {
  return msal.loginRedirect({ scopes: GRAPH_SCOPES, prompt: "select_account" });
}

export function signOut() {
  return msal.logoutRedirect({ account: msal.getActiveAccount() ?? undefined });
}

async function silent(scopes: string[], forceRefresh = false) {
  const account = msal.getActiveAccount();
  if (!account) throw new Error("Not signed in");
  try {
    return await msal.acquireTokenSilent({ scopes, account, forceRefresh });
  } catch (e) {
    // Session expired or consent needed: send the user back through Microsoft sign-in.
    if (e instanceof InteractionRequiredAuthError) await msal.acquireTokenRedirect({ scopes, account });
    throw e;
  }
}

/** Access token for Microsoft Graph (reading mail). */
export async function graphToken(): Promise<string> {
  return (await silent(GRAPH_SCOPES)).accessToken;
}

/** ID token proving who the user is to the Mail AI API (/v2/*). */
export async function idToken(forceRefresh = false): Promise<string> {
  return (await silent(GRAPH_SCOPES, forceRefresh)).idToken;
}
