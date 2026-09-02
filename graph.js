/* ==========================================================================
   ABC Application — Microsoft Graph / SharePoint integration layer
   Fill in the four CONFIG values below once you've done Steps 1-2 (create the
   SharePoint lists, register the Entra ID app). Everything else in this file
   works unmodified once CONFIG is correct.
   ========================================================================== */

const GRAPH_CONFIG = {
  clientId: "YOUR-AZURE-AD-APP-CLIENT-ID",
  tenantId: "YOUR-AZURE-AD-TENANT-ID",
  siteHostname: "yourtenant.sharepoint.com",   // no https://
  sitePath: "/sites/ABCApplication",           // your site's server-relative path
  redirectUri: window.location.origin + window.location.pathname
};

/* -------------------------------------------------------------------------
   Auth (MSAL.js — loaded via CDN in index.html, no build step needed)
   ------------------------------------------------------------------------- */
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: GRAPH_CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${GRAPH_CONFIG.tenantId}`,
    redirectUri: GRAPH_CONFIG.redirectUri
  },
  cache: { cacheLocation: "sessionStorage" }
});

const GRAPH_SCOPES = ["Sites.ReadWrite.All", "User.Read"]; // or Sites.Selected, see Step 2

async function signIn() {
  const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
  msalInstance.setActiveAccount(result.account);
  return result.account;
}

async function getGraphToken() {
  const account = msalInstance.getActiveAccount();
  if (!account) throw new Error("Not signed in");
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return result.accessToken;
  } catch (e) {
    const result = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return result.accessToken;
  }
}

async function graphFetch(path, options = {}) {
  const token = await getGraphToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) }
  });
  if (!res.ok) throw new Error(`Graph ${options.method || "GET"} ${path} failed: ${res.status} ${await res.text()}`);
  return res.status === 204 ? null : res.json();
}

/* -------------------------------------------------------------------------
   Site + list resolution (cached after first call)
   ------------------------------------------------------------------------- */
let _siteId = null;
async function getSiteId() {
  if (_siteId) return _siteId;
  const site = await graphFetch(`/sites/${GRAPH_CONFIG.siteHostname}:${GRAPH_CONFIG.sitePath}`);
  _siteId = site.id;
  return _siteId;
}

/* -------------------------------------------------------------------------
   Generic list CRUD (list name = the display name you gave it in SharePoint)
   ------------------------------------------------------------------------- */
async function graphListItems(listName, filterOData) {
  const siteId = await getSiteId();
  const filter = filterOData ? `&$filter=${encodeURIComponent(filterOData)}` : "";
  const data = await graphFetch(`/sites/${siteId}/lists/${listName}/items?expand=fields${filter}`);
  return data.value.map(item => ({ id: item.id, ...item.fields }));
}

async function graphCreateItem(listName, fields) {
  const siteId = await getSiteId();
  const item = await graphFetch(`/sites/${siteId}/lists/${listName}/items`, {
    method: "POST",
    body: JSON.stringify({ fields })
  });
  return { id: item.id, ...item.fields };
}

async function graphUpdateItem(listName, itemId, fields) {
  const siteId = await getSiteId();
  await graphFetch(`/sites/${siteId}/lists/${listName}/items/${itemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fields)
  });
}

async function graphDeleteItem(listName, itemId) {
  const siteId = await getSiteId();
  await graphFetch(`/sites/${siteId}/lists/${listName}/items/${itemId}`, { method: "DELETE" });
}
