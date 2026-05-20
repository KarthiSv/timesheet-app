/**
 * OneDrive / Microsoft Graph integration
 * Uses MSAL (Microsoft Authentication Library) with PKCE for browser-side OAuth.
 * No backend required. User provides their Azure App Client ID.
 */
import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

// ─── Constants ────────────────────────────────────────────────────────────────
export const DB_FOLDER   = "TimesheetManager";
export const DB_FILENAME = "TimesheetManager_DB.json";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Files.ReadWrite", "User.Read"];

// ─── MSAL instance (lazy-init) ────────────────────────────────────────────────
let _msal = null;

export async function initMsal(clientId) {
  if (!clientId) throw new Error("No Client ID provided");
  const config = {
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: window.location.origin,
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  };
  _msal = new PublicClientApplication(config);
  await _msal.initialize();
  return _msal;
}

export function getMsal() {
  if (!_msal) throw new Error("MSAL not initialised — call initMsal(clientId) first");
  return _msal;
}

// ─── Auth ─────────────────────────────────────────────────────────────────────
export async function signIn() {
  const msal = getMsal();
  const result = await msal.loginPopup({ scopes: SCOPES });
  return result.account;
}

export async function signOut() {
  const msal = getMsal();
  const account = msal.getAllAccounts()[0];
  if (account) await msal.logoutPopup({ account });
}

export function getAccount() {
  if (!_msal) return null;
  const accounts = _msal.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : null;
}

export function isSignedIn() {
  return getAccount() !== null;
}

async function getToken() {
  const msal = getMsal();
  const account = getAccount();
  if (!account) throw new Error("Not signed in to OneDrive");
  try {
    const result = await msal.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const result = await msal.acquireTokenPopup({ scopes: SCOPES });
      return result.accessToken;
    }
    throw e;
  }
}

// ─── Graph API helpers ────────────────────────────────────────────────────────
async function graphFetch(path, opts = {}) {
  const token = await getToken();
  const url   = path.startsWith("https://") ? path : GRAPH + path;
  const resp  = await fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(opts.body && typeof opts.body === "string" ? { "Content-Type": "application/json" } : {}),
      ...(opts.headers || {}),
    },
  });
  if (!resp.ok) {
    const msg = await resp.text().catch(() => resp.statusText);
    throw new Error(`Graph ${resp.status}: ${msg}`);
  }
  return resp;
}

// Get signed-in user profile
export async function getMe() {
  const resp = await graphFetch("/me?$select=displayName,mail,userPrincipalName");
  return resp.json();
}

// ─── Folder / file listing ────────────────────────────────────────────────────
export async function listFolder(folderId) {
  const path = (!folderId || folderId === "root")
    ? "/me/drive/root/children"
    : `/me/drive/items/${folderId}/children`;
  const resp = await graphFetch(
    `${path}?$select=id,name,size,lastModifiedDateTime,folder,file,parentReference&$orderby=name&$top=200`
  );
  const data = await resp.json();
  return data.value || [];
}

// ─── File operations ──────────────────────────────────────────────────────────
export async function downloadFile(itemId) {
  const resp = await graphFetch(`/me/drive/items/${itemId}/content`);
  return resp.arrayBuffer();
}

// Upload (< 4 MB) — pass ArrayBuffer or string
export async function uploadFile(folderIdOrPath, filename, content, contentType = "application/octet-stream") {
  const token = await getToken();
  let url;
  if (!folderIdOrPath || folderIdOrPath === "root") {
    url = `${GRAPH}/me/drive/root:/${encodeURIComponent(filename)}:/content`;
  } else if (folderIdOrPath.startsWith("/")) {
    // Path-based
    url = `${GRAPH}/me/drive/root:${folderIdOrPath}/${encodeURIComponent(filename)}:/content`;
  } else {
    // Item ID-based
    url = `${GRAPH}/me/drive/items/${folderIdOrPath}:/${encodeURIComponent(filename)}:/content`;
  }
  const resp = await fetch(url, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": contentType },
    body: content,
  });
  if (!resp.ok) {
    const msg = await resp.text().catch(() => resp.statusText);
    throw new Error(`Upload failed ${resp.status}: ${msg}`);
  }
  return resp.json();
}

// ─── App folder / database ────────────────────────────────────────────────────
async function ensureAppFolder() {
  try {
    const resp = await graphFetch(`/me/drive/root:/${DB_FOLDER}`);
    const item = await resp.json();
    return item.id;
  } catch {
    // Folder doesn't exist — create it
    const resp = await graphFetch("/me/drive/root/children", {
      method: "POST",
      body: JSON.stringify({ name: DB_FOLDER, folder: {}, "@microsoft.graph.conflictBehavior": "rename" }),
    });
    const item = await resp.json();
    return item.id;
  }
}

export async function loadDatabase() {
  const resp = await graphFetch(`/me/drive/root:/${DB_FOLDER}/${DB_FILENAME}:/content`);
  return resp.json();
}

export async function saveDatabase(data) {
  await ensureAppFolder();
  const json = JSON.stringify(data, null, 2);
  await uploadFile(`/${DB_FOLDER}`, DB_FILENAME, json, "application/json");
}

export async function databaseExists() {
  try {
    await graphFetch(`/me/drive/root:/${DB_FOLDER}/${DB_FILENAME}`);
    return true;
  } catch {
    return false;
  }
}
