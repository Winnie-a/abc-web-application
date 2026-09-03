/* ==========================================================================
   ABC Application — Microsoft Graph / SharePoint integration layer
   Fill in the four CONFIG values below once you've done Steps 1-2 (create the
   SharePoint lists, register the Entra ID app). Everything else in this file
   works unmodified once CONFIG is correct.
   ========================================================================== */

const GRAPH_CONFIG = {
  clientId: "af0925ee-e6d6-4477-88fe-926100eacf7f",
  tenantId: "58000a83-938f-4833-b798-e6b43eab9bb6",
  siteHostname: "rgbgames.sharepoint.com",   // no https://
  sitePath: "/sites/playground",             // your site's server-relative path
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
// MSAL Browser v3+ requires initialize() to resolve before any other MSAL
// API call — every function below that touches msalInstance awaits this.
const msalReady = msalInstance.initialize();

const GRAPH_SCOPES = ["Sites.ReadWrite.All", "User.Read", "User.Read.All", "GroupMember.Read.All"]; // or Sites.Selected, see Step 2
// User.Read.All and GroupMember.Read.All let the signed-in user's token read
// OTHER users' directory profiles and group memberships. A tenant admin must
// grant consent for both once, from this app's "API permissions" page in the
// Azure Portal — a regular user can't self-consent to them the way they can
// for Sites.ReadWrite.All/User.Read.

// Entra ID security groups used as role rosters for this app. Add one entry
// per role group (e.g. compliance committee, EXCO) as they're created —
// find each group's Object Id on its "Overview" page in the Azure Portal.
const GRAPH_GROUPS = {
  FCPA_DH: "0f4b2a0c-b245-4b04-82b6-b3d16bbb29c6",         // "FCPA DH" — Department Head roster
  HIGHER_MANAGEMENT: "a125a4bc-7645-4ef7-910c-2fbbb9261642" // "FCPA Higher Management" — skips DH, straight to CFO
};

function isGraphConnected() {
  // Synchronous by design (called from non-async code) — safe no-op if
  // msalReady hasn't resolved yet, rather than throwing.
  try {
    return !!msalInstance.getActiveAccount();
  } catch (e) {
    return false;
  }
}

async function signIn() {
  await msalReady;
  const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
  msalInstance.setActiveAccount(result.account);
  return result.account;
}

async function getGraphToken() {
  await msalReady;
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

/* -------------------------------------------------------------------------
   Corporate directory (Entra ID / Microsoft Graph)
   Maps a Graph user record into the shape EMPLOYEES/CURRENT_USER already use
   in data.js. Graph has no concept of this app's approval-routing fields
   (team, higherManagement, department-head assignment) — those still need
   to come from your own source (e.g. an ABC_EmployeeRouting SharePoint
   list keyed by employeeId), merged in after this mapping.
   ------------------------------------------------------------------------- */
const GRAPH_USER_SELECT = "id,displayName,mail,userPrincipalName,jobTitle,department,employeeId";

function mapGraphUser(u) {
  return {
    graphId: u.id,                        // Entra object id — use this (not employeeNo) for group-membership checks
    employeeNo: u.employeeId || u.id,     // falls back to the Graph object id if employeeId isn't synced from HR
    name: u.displayName,
    department: u.department || "",
    position: u.jobTitle || "",
    email: (u.mail || u.userPrincipalName || "").toLowerCase(),
    team: null,            // fill in from your routing source
    higherManagement: false // set via applyHigherManagementFlag() below
  };
}

// Marks `user` as Higher Management if they belong to GRAPH_GROUPS.HIGHER_MANAGEMENT
// (mirrors the isChuahFamily check in the existing Power Apps build, but reads
// live group membership instead of a hardcoded email list).
async function applyHigherManagementFlag(user) {
  const hmMembers = await graphGetGroupMembers(GRAPH_GROUPS.HIGHER_MANAGEMENT);
  user.higherManagement = hmMembers.some(m => m.graphId === user.graphId);
  return user;
}

async function graphGetMe() {
  const me = await graphFetch(`/me?$select=${GRAPH_USER_SELECT}`);
  return mapGraphUser(me);
}

async function graphListUsers() {
  const users = [];
  let path = `/users?$select=${GRAPH_USER_SELECT}&$top=999`;
  while (path) {
    const page = await graphFetch(path);
    users.push(...page.value.map(mapGraphUser));
    path = page["@odata.nextLink"] ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return users;
}

// Members of an Entra ID security group (e.g. GRAPH_GROUPS.FCPA_DH), mapped
// the same way as graphListUsers. Non-user members (e.g. nested groups) are
// skipped since they don't have the selected user fields.
async function graphGetGroupMembers(groupId) {
  const members = [];
  let path = `/groups/${groupId}/members?$select=${GRAPH_USER_SELECT}&$top=999`;
  while (path) {
    const page = await graphFetch(path);
    members.push(...page.value.filter(m => m["@odata.type"] === "#microsoft.graph.user" || m.userPrincipalName).map(mapGraphUser));
    path = page["@odata.nextLink"] ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return members;
}

/* -------------------------------------------------------------------------
   ABC Authority (SharePoint list) — the live role -> person roster (CFO,
   COO, GCOO, MD, Compliance Committee, ...), replacing hardcoded fixedName
   values in data.js. The exact internal SharePoint field name for the
   "Title Authority" column wasn't confirmed, so this tries a few plausible
   variants defensively — if none match, fix the fallback chain below to
   whatever the real internal name turns out to be (visible in a raw
   graphListItems("ABC Authority") result).
   ------------------------------------------------------------------------- */
function mapAuthorityRow(item) {
  return {
    name: item.Title,
    email: item.Email,
    position: item.Position,
    tag: item.TitleAuthority || item.Title_x0020_Authority || item.Title1 || ""
  };
}

async function graphGetAuthorityList() {
  const rows = await graphListItems("ABC Authority");
  return rows.map(mapAuthorityRow);
}
