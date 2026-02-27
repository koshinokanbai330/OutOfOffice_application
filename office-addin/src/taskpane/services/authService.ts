/* global Office */
import { PublicClientApplication, AccountInfo, InteractionRequiredAuthError } from "@azure/msal-browser";

const CLIENT_ID = process.env.REACT_APP_CLIENT_ID || "REPLACE_WITH_CLIENT_ID";
const TENANT_ID = process.env.REACT_APP_TENANT_ID || "common";

const SCOPES = [
  "User.Read",
  "MailboxSettings.ReadWrite",
  "Calendars.ReadWrite",
  "Files.ReadWrite",
  "offline_access",
];

let msalInstance: PublicClientApplication | null = null;

async function getMsalInstance(): Promise<PublicClientApplication> {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        redirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      },
    });
    await msalInstance.initialize();
  }
  return msalInstance;
}

export async function getAccessToken(): Promise<string> {
  // Try Office SSO first
  try {
    const ssoToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    return ssoToken;
  } catch {
    // Fall back to MSAL
    return getMsalToken();
  }
}

async function getMsalToken(): Promise<string> {
  const msal = await getMsalInstance();
  const accounts = msal.getAllAccounts();

  const request = {
    scopes: SCOPES,
    account: accounts[0] as AccountInfo | undefined,
  };

  try {
    const result = await msal.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const result = await msal.acquireTokenPopup(request);
      return result.accessToken;
    }
    throw err;
  }
}
