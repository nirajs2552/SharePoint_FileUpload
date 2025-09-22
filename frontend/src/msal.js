import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_AD_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_AD_TENANT_ID}`,
    redirectUri: window.location.origin
  },
  cache: { cacheLocation: "sessionStorage" }
};

export const msalInstance = new PublicClientApplication(msalConfig);
const SCOPES = (import.meta.env.VITE_GRAPH_SCOPES || "").split(" ").filter(Boolean);

export async function getAccount() {
  await msalInstance.initialize();
  return msalInstance.getAllAccounts()[0] || null;
}

export async function signIn() {
  await msalInstance.initialize();
  const account = await getAccount();
  if (account) return account;
  const login = await msalInstance.loginPopup({ scopes: SCOPES });
  return login.account;
}

export async function getAccessToken() {
  await msalInstance.initialize();
  let account = await getAccount();
  if (!account) {
    const login = await msalInstance.loginPopup({ scopes: SCOPES });
    account = login.account;
  }
  try {
    const { accessToken } = await msalInstance.acquireTokenSilent({ account, scopes: SCOPES });
    return accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const { accessToken } = await msalInstance.acquireTokenPopup({ account, scopes: SCOPES });
      return accessToken;
    }
    throw e;
  }
}

export async function signOut() {
  await msalInstance.initialize();
  const account = await getAccount();
  if (!account) return;
  // Use popup so your SPA stays on the same page during local dev
  await msalInstance.logoutPopup({ account });
}
