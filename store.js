/* ==========================================================================
   ABC Application — data store
   localStorage-backed model of pre-approvals & claims, standing in for the
   SharePoint list / Power Automate backend described in the release plan.
   ========================================================================== */

const STORAGE_KEY = "abc_app_state_v2";

function uid() {
  return Math.random().toString(36).slice(2, 10);
}
function genRefNo(country) {
  // Matches the original Power Apps build's formula:
  //   "ABC-" & Office365Users.MyProfile().Country & "-" & Left(GUID(), 8)
  // i.e. the reference number is tagged with the SUBMITTER's own Entra ID
  // "Country" profile field (whoever is signed in and clicks Submit) — not
  // the requestor being submitted for, and not the recipient's country.
  // Falls back to "MY" if the signed-in profile has no Country set, so we
  // never emit a double-dash "ABC--xxxxx" ref no.
  const tag = (country || "").toString().trim() || "MY";
  return "ABC-" + tag + "-" + uid();
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function fmtDate(iso) {
  if (!iso) return "-";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/* The six-month recipient history used to be a seeded mock generated here
   (sixMonthHistory). It now comes from the real SharePoint lists — see
   graphGetRecipientSpendHistory() in graph.js and App.recipientHistory() in
   app.js. */

function rowTotal(row) {
  return row.gifts + row.meals + row.travel + row.entertainment + row.others;
}

/* --------------------------------------------------------------------------
   Approval chain resolution
   -------------------------------------------------------------------------- */

/* Build a concrete chain for a record from a stage-definition list (see
   preApprovalChainDef / claimChainDef): resolves each stage's approver name
   — either a fixed name or one looked up via resolveMap (dh/hod/shod) — and
   sets the first stage Pending, the rest Waiting. Every stage in the given
   list is required, in order — any amount-based gating (Claim flow only;
   see claimChainDef) has already been decided before this list was built. */
function buildChain(stageDefs, resolveMap) {
  const stages = stageDefs.map(def => ({
    title: def.title,
    role: def.role,
    name: def.resolve ? (resolveMap[def.resolve] || "") : def.fixedName,
    applicable: true,
    status: "waiting", // waiting | pending | approved | rejected
    date: null,
    comments: "",
    rejectReason: "",
    // Written to the SharePoint Gate<N><Stage>Email / Position columns. Left
    // blank (never guessed) when the approver isn't in the ABC Authority
    // roster or the Entra directory.
    email: lookupApproverEmail(def.resolve ? (resolveMap[def.resolve] || "") : def.fixedName, def.fixedName ? def.title : ""),
    position: lookupApproverPosition(def.resolve ? (resolveMap[def.resolve] || "") : def.fixedName) || def.role
  }));
  if (stages.length) stages[0].status = "pending";
  return stages;
}

/* Real (non-guessed) approver contact details: the admin-managed "ABC
   Authority" roster first (App.state.approvers), then the Entra ID directory
   (App.state.directory) matched on display name. Both are only populated
   after sign-in, so this returns "" for seed/demo data. */
// Names compared without "(nickname)" suffixes and extra spaces, e.g. the
// directory's "Liew Yung Kuan (YK)" matches the stage name "Liew Yung Kuan".
function _normName(s) {
  return String(s || "").toLowerCase().replace(/\(.*?\)/g, "").replace(/\s+/g, " ").trim();
}
function _approverRows(name, stageTitle) {
  if (typeof App === "undefined" || !App.state) return [];
  const n = _normName(name);
  const hits = [];
  if (n) {
    (App.state.approvers || []).forEach(r => { if (_normName(r.name) === n) hits.push(r); });
    (App.state.directory || []).forEach(u => { if (_normName(u.name) === n) hits.push(u); });
  }
  // fixed roles (CFO / GCOO / MD): the ABC Authority roster row for that role
  if (!hits.length && stageTitle) {
    (App.state.approvers || []).forEach(r => { if (String(r.role || "").trim().toLowerCase() === String(stageTitle).trim().toLowerCase()) hits.push(r); });
  }
  return hits;
}
function lookupApproverEmail(name, stageTitle) {
  const hit = _approverRows(name, stageTitle).find(r => r.email && String(r.email).trim());
  return hit ? String(hit.email).trim() : "";
}
function lookupApproverPosition(name) {
  const hit = _approverRows(name).find(r => r.role || r.position);
  return hit ? (hit.role || hit.position) : "";
}

function chainOverallStatus(stages, openLabel, closedLabel) {
  if (stages.some(s => s.status === "rejected")) return "Rejected";
  if (stages.length && stages.every(s => s.status === "approved")) return closedLabel;
  return openLabel;
}

/* Advance chain after an approve/reject action on stageIndex. Mutates stages. */
function advanceChain(stages, stageIndex, action, comments, rejectReason) {
  const stage = stages[stageIndex];
  stage.date = todayISO();
  stage.comments = comments || "";
  if (action === "reject") {
    stage.status = "rejected";
    stage.rejectReason = rejectReason || "";
    return;
  }
  stage.status = "approved";
  if (stages[stageIndex + 1]) stages[stageIndex + 1].status = "pending";
}

/* Resolve the {dh, hod, shod} name map a chain needs for a given requestor
   + chosen department head. */
function chainResolveMap(requestor, departmentHead) {
  const emp = employeeByNo(requestor.employeeNo);
  const line = orgLineFor(emp && emp.team);
  return { dh: departmentHead, hod: line.hod, shod: line.shod };
}

/* --------------------------------------------------------------------------
   Seed data + persistence
   -------------------------------------------------------------------------- */

function computePreApprovalTotalUSD(rec) {
  const a = rec.amounts;
  const sum = AMOUNT_FIELDS.reduce((s, f) => s + (Number(a[f.key]) || 0), 0) + (Number(a.othersAmount) || 0);
  return toUSD(sum, rec.currency);
}

function newBlankPreApproval() {
  return {
    id: uid(),
    refNo: genRefNo(CURRENT_USER.country),
    submittedBy: CURRENT_USER.name,
    requestor: { name: CURRENT_USER.name, employeeNo: CURRENT_USER.employeeNo, department: CURRENT_USER.department, position: CURRENT_USER.position },
    recipients: [],
    transactionTypes: [],
    description: "",
    currency: "",
    amounts: { gifts: 0, meals: 0, entertainment: 0, airfare: 0, transportation: 0, hotel: 0, othersLabel: "", othersAmount: 0 },
    paymentTo: "",
    remarks: "",
    departmentHead: "",
    dateSubmitted: null,
    approvals: [],
    status: "Draft",
    claim: null
  };
}

function buildSeedPreApproval({ requestor, recipients, currency, amounts, dh, daysAgo, forceStatus, withClaim, claimClosed }) {
  const rec = newBlankPreApproval();
  rec.requestor = requestor;
  rec.recipients = recipients;
  rec.transactionTypes = ["Gifts", "Meals", "Travel"];
  rec.description = "Client relationship dinner and courtesy gifts during Q3 site visit.";
  rec.currency = currency;
  rec.amounts = amounts;
  rec.paymentTo = "Paid directly to venue / recipient by requestor, claimed back via ABC Claim.";
  rec.remarks = "Standard courtesy engagement, within department budget.";
  rec.departmentHead = dh;
  const d = new Date(); d.setDate(d.getDate() - daysAgo);
  rec.dateSubmitted = d.toISOString().slice(0, 10);

  rec.approvals = buildChain(preApprovalChainDef(isHigherManagement(rec.requestor)), chainResolveMap(rec.requestor, dh));

  if (forceStatus === "approved" || forceStatus === "closed") {
    rec.approvals.forEach(s => { s.status = "approved"; s.date = rec.dateSubmitted; s.comments = "Please proceed. Thanks."; });
    rec.status = "Approved";
  } else {
    rec.status = "Pending";
  }

  if (withClaim) {
    rec.claim = newBlankClaim(rec, forceStatus === "closed");
    if (forceStatus === "closed" && claimClosed) {
      rec.claim.approvals.forEach(s => { s.status = "approved"; s.date = rec.dateSubmitted; });
      rec.claim.status = "Closed";
    }
  }
  return rec;
}

function newBlankClaim(preApproval, prefill) {
  const claim = {
    company: prefill ? "RGB Sdn Bhd" : "",
    costBearBy: prefill ? "RGB Sdn Bhd" : "",
    date: prefill ? preApproval.dateSubmitted : "",
    conversionRate: prefill ? 4.2 : "",
    // Currency the claim's line-item amounts are entered in — same as the
    // pre-approval's currency, as in the Power App ("Amount in <Currency>").
    // Seed claims below are MYR demo data, so they pin MYR to keep their totals.
    currency: prefill ? "MYR" : (preApproval.currency || ""),
    purposeOfClaim: prefill ? preApproval.description : "",
    lineItems: [],
    attachments: [],
    /* Recipients always carry over from the original pre-approval — the
       manual describes them as pre-populated, with "Add User" / "Click here"
       only used to add extras or ones missing from the drop-down. */
    register: {
      month: "", year: preApproval.dateSubmitted ? Number(preApproval.dateSubmitted.slice(0, 4)) : 2026,
      recipients: preApproval.recipients.map(r => ({
        date: preApproval.dateSubmitted, name: r.name, company: r.company, others: 0, isOfficial: r.isOfficial || "No"
      }))
    },
    approvals: [],
    status: "Open"
  };
  if (prefill) {
    claim.lineItems = preApproval.recipients.slice(0, 1).map((r, i) => ({
      id: uid(), date: preApproval.dateSubmitted, description: "Dinner with " + r.name,
      hasReceipt: "Yes", transactionType: "Meals", paymentMethod: "Cash",
      amount: 620, purpose: "Client engagement", attachment: "receipt-0" + (i + 1) + ".pdf"
    }));
    claim.register.month = "June";
    claim.approvals = buildChain(claimChainDef(isHigherManagement(preApproval.requestor), claimActualUSD(claim)), chainResolveMap(preApproval.requestor, preApproval.departmentHead));
  }
  return claim;
}

/* Claim line-item money — ported from the Power App's ClaimForm screen
   (TotalAmountClaimFormListConverted), which converts each line on its own:
     - the line's rate is the claim's header Conversion Rate for Cash (and any
       other non-card method), or the line's own bank rate for Credit Card;
     - rates are entered as "USD to <currency>" (e.g. 4.2 for MYR), so USD =
       amount / rate for rates >= 1 and amount * rate for rates < 1 — except
       EUR / GBP, which are entered the other way round (1.09 meaning
       EUR->USD), so the two branches flip;
     - a rate of 0 contributes 0 USD.
   Older saved lines used amountMYR (always MYR) — lineAmount()/claimCurrency()
   keep those readable. */
function claimCurrency(claim) {
  return claim.currency || "MYR";
}
function lineAmount(li) {
  return Number(li.amount != null ? li.amount : li.amountMYR) || 0;
}
function lineRate(li, claim) {
  return li.paymentMethod === "Credit Card" ? (Number(li.rate) || 0) : (Number(claim.conversionRate) || 0);
}
function lineAmountUSD(amount, rate, currency) {
  if (!(rate > 0)) return 0;
  const cur = (currency || "").toUpperCase();
  const flipped = cur === "EUR" || cur === "EURO" || cur === "GBP";
  const multiply = flipped ? rate >= 1 : rate < 1;
  return multiply ? amount * rate : amount / rate;
}
function claimActualLocal(claim) {
  return claim.lineItems.reduce((s, li) => s + lineAmount(li), 0);
}
function claimActualUSD(claim) {
  const cur = claimCurrency(claim);
  return claim.lineItems.reduce((s, li) => s + lineAmountUSD(lineAmount(li), lineRate(li, claim), cur), 0);
}

function seedState() {
  const dh1 = defaultDHForEmployee(CURRENT_USER) || "Chow Bong Weng";
  const otherEmp = EMPLOYEES[1];
  const dh2 = defaultDHForEmployee(otherEmp) || "Lim Cheng Soon";

  const pending = buildSeedPreApproval({
    requestor: { name: CURRENT_USER.name, employeeNo: CURRENT_USER.employeeNo, department: CURRENT_USER.department, position: CURRENT_USER.position },
    recipients: [
      { name: "Lee Seong Gi", company: "D'heights Resort & Casino Clark Philippines", position: "President", relationship: "Customer", isOfficial: "No" }
    ],
    currency: "USD",
    amounts: { gifts: 0, meals: 320, entertainment: 0, airfare: 0, transportation: 0, hotel: 0, othersLabel: "", othersAmount: 0 },
    dh: dh1, daysAgo: 3, forceStatus: "pending", withClaim: false
  });

  const approvedOpenClaim = buildSeedPreApproval({
    requestor: { name: CURRENT_USER.name, employeeNo: CURRENT_USER.employeeNo, department: CURRENT_USER.department, position: CURRENT_USER.position },
    recipients: [
      { name: "Lee Seong Gi", company: "D'heights Resort & Casino Clark Philippines", position: "President", relationship: "Customer", isOfficial: "No" },
      { name: "Lee Sung Gyun", company: "D'heights Resort & Casino Clark Philippines", position: "Finance Manager", relationship: "Customer", isOfficial: "No" }
    ],
    currency: "USD",
    amounts: { gifts: 0, meals: 620, entertainment: 0, airfare: 1275, transportation: 0, hotel: 0, othersLabel: "", othersAmount: 0 },
    dh: dh1, daysAgo: 14, forceStatus: "approved", withClaim: true
  });

  const closedHistory1 = buildSeedPreApproval({
    requestor: { name: otherEmp.name, employeeNo: otherEmp.employeeNo, department: otherEmp.department, position: otherEmp.position },
    recipients: [{ name: "Ms Gladys Lei", company: "IGT", position: "Regional Manager", relationship: "Business Partner / JV Partner", isOfficial: "No" }],
    currency: "USD",
    amounts: { gifts: 150, meals: 0, entertainment: 0, airfare: 0, transportation: 0, hotel: 0, othersLabel: "", othersAmount: 0 },
    dh: dh2, daysAgo: 85, forceStatus: "closed", withClaim: true, claimClosed: true
  });

  const closedHistory2 = buildSeedPreApproval({
    requestor: { name: CURRENT_USER.name, employeeNo: CURRENT_USER.employeeNo, department: CURRENT_USER.department, position: CURRENT_USER.position },
    recipients: [{ name: "Charles Seo", company: "The Grand Ho Tram Strip Casino", position: "General Manager", relationship: "Customer", isOfficial: "No" }],
    currency: "USD",
    amounts: { gifts: 0, meals: 0, entertainment: 0, airfare: 2200, transportation: 340, hotel: 980, othersLabel: "", othersAmount: 0 },
    dh: dh1, daysAgo: 60, forceStatus: "closed", withClaim: true, claimClosed: true
  });

  return { preApprovals: [pending, approvedOpenClaim, closedHistory1, closedHistory2] };
}

/* First stage currently awaiting action — who a real email integration
   would notify next. */
function pendingNotifyTarget(stages) {
  const s = (stages || []).find(st => st.status === "pending");
  if (!s || !s.name) return null;
  return { name: s.name, title: s.title, email: Store.getApproverEmail(s.name) };
}

/* --------------------------------------------------------------------------
   SharePoint push — ABC Pre-Approval + ABC Recipient Final Expenses
   Internal SharePoint column names below are a BEST GUESS (display name
   with spaces stripped, e.g. "Others Type" -> "OthersType"), except
   FCPA No -> Title, which is SharePoint's mandatory default column,
   renamed. None of this has been verified against the real list yet —
   verify once a real sign-in can actually reach it, and fix any names
   Graph rejects (the error will name the offending field).
   -------------------------------------------------------------------------- */

// Approval stage title -> the "Gate 0"/"Gate 1" column-name suffix used in
// SharePoint (Compliance Committee 1/2 shorten to Comp/Comp2; everything
// else matches the stage title as-is).
function gateColumnKey(stageTitle) {
  const map = { "Compliance Committee 1": "Comp", "Compliance Committee 2": "Comp2" };
  return map[stageTitle] || stageTitle.replace(/\s+/g, "");
}

// The per-stage "...ApprovalStatus" columns are Choice fields that only accept
// N/A / Pending / Approved / Rejected (verified 2026-10-07 on the live list),
// so the app's own lowercase statuses are translated here. A "waiting" stage
// returns null: it hasn't started, so the column is left blank.
function spStageStatus(status) {
  return { pending: "Pending", approved: "Approved", rejected: "Rejected" }[status] || null;
}
// The overall "Gate0" Choice column: Closed / Pending / Rejected / Cancelled /
// Overdue. Same wording the Power Automate flows write (Closed = fully approved).
function spOverallStatus(stages) {
  if (stages.some(s => s.status === "rejected")) return "Rejected";
  if (stages.length && stages.every(s => s.status === "approved")) return "Closed";
  return "Pending";
}

function gateFields(gatePrefix, stage) {
  const key = gatePrefix + gateColumnKey(stage.title);
  const fields = {
    [key]: stage.name || "",
    [key + "Email"]: stage.email || "",
    [key + "Position"]: stage.position || "",
    [key + "Comments"]: stage.comments || "",
    [key + "RejectReason"]: stage.rejectReason || ""
  };
  const status = spStageStatus(stage.status);
  if (status) fields[key + "ApprovalStatus"] = status;
  if (stage.date) fields[key + "ApprovalDate"] = stage.date;
  return fields;
}

// SharePoint list item id of this request's ABC Pre-Approval row. Uses the id
// remembered at submit time; otherwise looks the row up by its FCPA No (the
// ABC reference number) and remembers it. Resolves to null if there is no row.
async function findPreApprovalRowId(rec) {
  if (rec.spId) return rec.spId;
  const siteId = await getSiteId();
  const filter = encodeURIComponent(`fields/FCPANo eq '${String(rec.refNo).replace(/'/g, "''")}'`);
  const data = await graphFetch(
    `/sites/${siteId}/lists/${resolveListRef("ABC Pre-Approval")}/items?$filter=${filter}&$select=id`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  const row = data.value && data.value[0];
  if (row) rec.spId = row.id;
  return row ? row.id : null;
}

// Saves ONE approve/reject decision on a Pre-Approval stage to SharePoint —
// the stage's own status/date/comments/reject reason, the next stage flipping
// to Pending, and the overall Gate0 status. Only for requests submitted from
// this app (rec.live) — seeded demo records are never written. If the request
// has no SharePoint row yet (e.g. its first save failed), the whole request is
// created with its current state instead, so nothing is lost.
async function savePreApprovalDecision(rec, stageIndex) {
  const stage = rec.approvals[stageIndex];
  const itemId = await findPreApprovalRowId(rec);
  if (!itemId) { await pushPreApprovalToSharePoint(rec); return; }
  const key = "Gate0" + gateColumnKey(stage.title);
  const fields = {
    Gate0: spOverallStatus(rec.approvals),
    [key + "ApprovalStatus"]: spStageStatus(stage.status),
    [key + "ApprovalDate"]: stage.date,
    [key + "Comments"]: stage.comments || "",
    [key + "RejectReason"]: stage.rejectReason || ""
  };
  const next = rec.approvals[stageIndex + 1];
  if (stage.status === "approved" && next) fields["Gate0" + gateColumnKey(next.title) + "ApprovalStatus"] = "Pending";
  await graphUpdateItem("ABC Pre-Approval", itemId, fields);
}

async function pushPreApprovalToSharePoint(rec) {
  const a = rec.amounts;
  const fields = {
    FCPANo: rec.refNo,
    NameofRequestor: rec.requestor.name,
    // NOTE: no "SubmittedBy" here — that column doesn't exist on the list (checked
    // 2026-10-07), and Graph rejects the whole create on an unknown field. The
    // submitter is already recorded by SharePoint's own "Created By".
    EmployeeNo: rec.requestor.employeeNo,
    Department: rec.requestor.department,
    Position: rec.requestor.position,
    Email: rec.requestor.email || "",
    ProposedTransaction: rec.transactionTypes || [], // MultiChoice field — array, not a joined string
    "ProposedTransaction@odata.type": "Collection(Edm.String)", // Graph rejects a multi-choice array (400 "Invalid request") without this type hint
    ProposedTransactionDescription: rec.description || "",
    Currency: rec.currency,
    Gifts: a.gifts, Meals: a.meals, Entertainment: a.entertainment, Airfare: a.airfare,
    Transportation: a.transportation, Hotel: a.hotel,
    OthersType: a.othersLabel, OthersAmount: a.othersAmount,
    PaymentInfo: rec.paymentTo, Remarks: rec.remarks,
    Gate0: spOverallStatus(rec.approvals),
    IsCancelled: false
  };
  rec.approvals.forEach(stage => Object.assign(fields, gateFields("Gate0", stage)));

  const item = await graphCreateItem("ABC Pre-Approval", fields);
  rec.spId = item.id;

  // Recipient rows (one per recipient). "ABCNo" is a lookup into this
  // request's ABC Pre-Approval row and "RecipientID" a lookup into the FCPA
  // Customer list; lookups are written as <Column>LookupId = the target row's
  // numeric id. Recipients picked from the live FCPA Customer list carry that
  // id as r.recipientId; one that isn't from the directory has none, so the
  // RecipientID link is simply left blank for it. (The old "FCPANo" lookup on
  // this list still points at FCPA Pre-Approval and is not written.)
  for (const r of rec.recipients) {
    const row = {
      Title: r.name || "",
      ABCNoLookupId: Number(item.id),
      Gifts: 0, Meals: 0, Entertainment: 0, Travel: 0, Others: 0
    };
    const rid = Number(r.recipientId);
    if (Number.isFinite(rid) && rid > 0) row.RecipientIDLookupId = rid;
    await graphCreateItem("ABC Recipient Final Expenses", row);
  }
}

/* --------------------------------------------------------------------------
   SharePoint push — ABC Claim Form
   One row per claim, created when the claim is submitted and updated on every
   approve/reject. Column facts below were read from the live list on
   2026-10-07:
     - Title holds the ABC reference number — that is how a claim row is found
       again. The "ABCNo" lookup (added 2026-10-07) links it to its ABC
       Pre-Approval row. The older "FCPANo" lookup is deliberately NOT written:
       it still points at the FCPA Pre-Approval list (d2160a4e…), so an ABC
       item id there would link the claim to an unrelated FCPA row.
     - Gate1 (overall) choices: Open / Pending / Closed / Cancelled / Rejected
       ("Rejected" added to the list 2026-10-07 to match the FCPA Claim Form's
       choices). A fully approved claim is Closed and a rejected one is Rejected,
       so the list matches what the web app shows. (The Power Automate flows
       instead reset Gate1 to Open on a rejection so the claim can be resubmitted;
       the web app has no resubmit step.)
     - Stage columns: HOD, SHOD, DH, CFO, COO, GCOO have clean names
       (Gate1<Stage>..., Gate1<Stage>Email...). MD and the 7 EXCO seats were
       created later and carry SharePoint's mangled, truncated internal names,
       listed explicitly in CLAIM_MD_COLS / claimExcoCols() below.
   -------------------------------------------------------------------------- */
const CLAIM_MD_COLS = {
  name: "Gate_x0020_1_x0020_MD", email: "Gate_x0020_1_x0020_MD_x0020_Emai", position: "Gate_x0020_1_x0020_MD_x0020_Posi",
  status: "Gate_x0020_1_x0020_MD_x0020_Appr", date: "Gate_x0020_1_x0020_MD_x0020_Appr0",
  comments: "Gate_x0020_1_x0020_MD_x0020_Comm", reject: "Gate_x0020_1_x0020_MD_x0020_Reje"
};
function claimExcoCols(seat) {
  const p = "Gate_x0020_1_x0020_EXCO_x0020_" + seat;
  return { name: p, email: p + "_", position: p + "_0", status: p + "_1", date: p + "_2", comments: p + "_3", reject: p + "_4" };
}
// Column names for the stage at stages[idx]. EXCO stages map to seats 1..7 in
// chain order (SVP first ... MD last), the same order the list's seats use.
function claimStageColumns(stages, idx) {
  const title = stages[idx].title;
  if (title === "MD") return CLAIM_MD_COLS;
  if (/^EXCO - /.test(title)) {
    const seat = stages.slice(0, idx + 1).filter(s => /^EXCO - /.test(s.title)).length;
    return claimExcoCols(seat);
  }
  const p = "Gate1" + title.replace(/\s+/g, "");
  return { name: p, email: p + "Email", position: p + "Position", status: p + "ApprovalStatus", date: p + "ApprovalDate", comments: p + "Comments", reject: p + "RejectReason" };
}
function spClaimOverallStatus(stages) {
  if (stages.some(s => s.status === "rejected")) return "Rejected";
  if (stages.length && stages.every(s => s.status === "approved")) return "Closed";
  return "Pending";
}
function claimGateFields(stages, idx) {
  const stage = stages[idx], c = claimStageColumns(stages, idx), f = {};
  f[c.name] = stage.name || "";
  f[c.email] = stage.email || "";
  f[c.position] = stage.position || "";
  f[c.comments] = stage.comments || "";
  f[c.reject] = stage.rejectReason || "";
  const status = spStageStatus(stage.status);
  if (status) f[c.status] = status;
  if (stage.date) f[c.date] = stage.date;
  return f;
}

// The claim row's SharePoint id lives on the request itself (rec.claimSpId):
// the row can exist (Gate 1 = Open) before the app has a local claim object.
// rec.claim.spId is the older location, still honoured.
function setClaimRowId(rec, id) {
  rec.claimSpId = id;
  if (rec.claim) rec.claim.spId = id;
}
async function findClaimRowId(rec) {
  const known = rec.claimSpId || (rec.claim && rec.claim.spId);
  if (known) return known;
  const siteId = await getSiteId();
  const filter = encodeURIComponent(`fields/Title eq '${String(rec.refNo).replace(/'/g, "''")}'`);
  const data = await graphFetch(
    `/sites/${siteId}/lists/${resolveListRef("ABC Claim Form")}/items?$filter=${filter}&$select=id`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  const row = data.value && data.value[0];
  if (row) setClaimRowId(rec, row.id);
  return row ? row.id : null;
}

// Same as the flows' "Create claim form" step: when the LAST Pre-Approval stage
// approves (Gate 0 = Closed), the claim row is created with Gate 1 = "Open",
// ready for the claim to be filled in and submitted. Does nothing if the row
// already exists.
async function createOpenClaimRow(rec) {
  if (await findClaimRowId(rec)) return;
  const fields = { Title: rec.refNo, Gate1: "Open", IsCancelled: false };
  const preId = Number(await findPreApprovalRowId(rec));
  if (Number.isFinite(preId) && preId > 0) fields.ABCNoLookupId = preId;
  const item = await graphCreateItem("ABC Claim Form", fields);
  setClaimRowId(rec, item.id);
}

// Claim submit: fills in the claim's ABC Claim Form row (header + every
// approval stage in its current state) and moves Gate 1 to Pending. Updates the
// "Open" row created by createOpenClaimRow() — or creates the row if there
// isn't one yet. Line items and the recipient register are not pushed yet.
async function pushClaimToSharePoint(rec) {
  const c = rec.claim;
  const fields = {
    Title: rec.refNo,
    Gate1: spClaimOverallStatus(c.approvals),
    SubmissionDate: c.submittedDate || todayISO(),
    ConversionRate: Number(c.conversionRate) || 0,
    Purposeofclaim: c.purposeOfClaim || "",
    Month: c.register && c.register.month || "",
    Year: c.register && Number(c.register.year) || 0,
    Total: claimActualLocal(c), // total in the claim's own currency
    CostBearCompany: c.costBearBy || "",
    Company: c.company || "",
    IsCancelled: false
  };
  c.approvals.forEach((_, i) => Object.assign(fields, claimGateFields(c.approvals, i)));
  // "ABCNo" is a lookup into this request's ABC Pre-Approval row (the old
  // "FCPANo" lookup still points at FCPA Pre-Approval, so it stays empty).
  const preId = Number(await findPreApprovalRowId(rec));
  if (Number.isFinite(preId) && preId > 0) fields.ABCNoLookupId = preId;
  const existingId = await findClaimRowId(rec);
  if (existingId) {
    await graphUpdateItem("ABC Claim Form", existingId, fields);
  } else {
    const item = await graphCreateItem("ABC Claim Form", fields);
    setClaimRowId(rec, item.id);
  }
}

// Saves ONE approve/reject on a Claim stage — same shape as
// savePreApprovalDecision(). If the claim has no row yet, the whole claim is
// created with its current state instead.
async function saveClaimDecision(rec, stageIndex) {
  const stages = rec.claim.approvals, stage = stages[stageIndex];
  const itemId = await findClaimRowId(rec);
  if (!itemId) { await pushClaimToSharePoint(rec); return; }
  const c = claimStageColumns(stages, stageIndex);
  const overall = spClaimOverallStatus(stages);
  const fields = {
    Gate1: overall,
    [c.status]: spStageStatus(stage.status),
    [c.date]: stage.date,
    [c.comments]: stage.comments || "",
    [c.reject]: stage.rejectReason || ""
  };
  const next = stages[stageIndex + 1];
  if (stage.status === "approved" && next) fields[claimStageColumns(stages, stageIndex + 1).status] = "Pending";
  if (overall === "Closed") fields.CompleteApprovalDate = stage.date;
  await graphUpdateItem("ABC Claim Form", itemId, fields);
}

/* --------------------------------------------------------------------------
   Recipient Final Expenses — per-recipient spend, the data behind the
   six-month history. Worked out from the FCPA lists (2026-10-07): the Power
   App takes each claim line in USD (amount ÷ conversion rate, the same
   conversion as the claim total), puts it in the column for its transaction
   type (Gifts / Meals / Entertainment / Travel / Others) and divides it
   EQUALLY between the recipients on the claim. Done when the claim is
   submitted, so the figures are in place before approvals finish; the history
   only counts them once the claim is Closed.
   -------------------------------------------------------------------------- */
const EXPENSE_CATEGORIES = ["Gifts", "Meals", "Entertainment", "Travel", "Others"];
const RECIPIENT_EXPENSE_LIST = "ABC Recipient Final Expenses";

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

// Claim lines -> USD total per expense column. A transaction type that isn't
// one of the columns counts as Others.
function claimCategoryTotalsUSD(claim) {
  const cur = claimCurrency(claim);
  const totals = { Gifts: 0, Meals: 0, Entertainment: 0, Travel: 0, Others: 0 };
  (claim.lineItems || []).forEach(li => {
    const key = EXPENSE_CATEGORIES.includes(li.transactionType) ? li.transactionType : "Others";
    totals[key] += lineAmountUSD(lineAmount(li), lineRate(li, claim), cur);
  });
  EXPENSE_CATEGORIES.forEach(k => { totals[k] = round2(totals[k]); });
  return totals;
}

// total split into n equal shares to the cent; any leftover cents go to the
// first recipient so the shares always add back up to the total exactly.
function splitEvenly(total, n) {
  const cents = Math.round(total * 100);
  const base = Math.floor(cents / n);
  const shares = Array(n).fill(base / 100);
  shares[0] = (base + (cents - base * n)) / 100;
  return shares;
}

function recipientIdFor(rec, person) {
  const name = String(person.name || "").trim().toLowerCase();
  const fromRequest = (rec.recipients || []).find(r => String(r.name || "").trim().toLowerCase() === name && r.recipientId);
  if (fromRequest) return Number(fromRequest.recipientId);
  const dir = (typeof App !== "undefined" && App.state && App.state.recipients) || [];
  const hit = dir.find(r => String(r.name || "").trim().toLowerCase() === name && (!person.company || r.company === person.company));
  return hit ? Number(hit.id) : 0;
}

// This request's rows in the list (they were created at Pre-Approval submit).
async function listRecipientExpenseRows(preId) {
  try {
    const siteId = await getSiteId();
    const filter = encodeURIComponent(`fields/ABCNoLookupId eq ${preId}`);
    const data = await graphFetch(
      `/sites/${siteId}/lists/${resolveListRef(RECIPIENT_EXPENSE_LIST)}/items?$filter=${filter}&$expand=fields($select=Title,ABCNoLookupId,RecipientIDLookupId,ClaimRegisterDate)`,
      { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
    );
    return (data.value || []).map(i => ({ id: i.id, ...i.fields }));
  } catch (e) {
    // server-side filter refused -> read the (small) list and filter here
    const all = await graphListAllItems(RECIPIENT_EXPENSE_LIST, ["Title", "ABCNoLookupId", "RecipientIDLookupId", "ClaimRegisterDate"]);
    return all.filter(r => Number(r.ABCNoLookupId) === Number(preId));
  }
}

// Fills in each claim-register recipient's share. A register recipient with no
// row yet (added at claim time) gets one; recipients removed from the register
// keep their row untouched. A per-recipient "Others" amount typed in the
// register is added to that recipient's Others.
async function saveRecipientExpenses(rec) {
  const claim = rec.claim;
  const people = (claim && claim.register && claim.register.recipients) || [];
  if (!people.length) return;
  const preId = Number(await findPreApprovalRowId(rec));
  if (!(preId > 0)) return;
  const totals = claimCategoryTotalsUSD(claim);
  const shares = {};
  EXPENSE_CATEGORIES.forEach(k => { shares[k] = splitEvenly(totals[k], people.length); });
  const rows = await listRecipientExpenseRows(preId);
  const used = new Set();
  for (let i = 0; i < people.length; i++) {
    const p = people[i];
    const name = String(p.name || "").trim().toLowerCase();
    const row = rows.find(r => !used.has(r.id) && String(r.Title || "").trim().toLowerCase() === name);
    if (row) used.add(row.id);
    const fields = {};
    EXPENSE_CATEGORIES.forEach(k => { fields[k] = shares[k][i]; });
    fields.Others = round2(fields.Others + (Number(p.others) || 0));
    if (p.date) fields.ClaimRegisterDate = p.date;
    if (row) {
      await graphUpdateItem(RECIPIENT_EXPENSE_LIST, row.id, fields);
    } else {
      const rid = recipientIdFor(rec, p);
      await graphCreateItem(RECIPIENT_EXPENSE_LIST, { Title: p.name || "", ABCNoLookupId: preId, ...(rid ? { RecipientIDLookupId: rid } : {}), ...fields });
    }
  }
}

/* --------------------------------------------------------------------------
   Claim line items -> the ABC Invoice list (one row per line), so an approver
   on another computer can see what is being claimed. "ABCClaimID" is a lookup
   into ABC Claim Form (the older "FCPAClaimID" lookup on this list points at
   the FCPA Claim Form, so it is not written). Same columns as FCPA Invoice.
   -------------------------------------------------------------------------- */
const INVOICE_LIST = "ABC Invoice";

async function listInvoiceRows(claimRowId) {
  try {
    const siteId = await getSiteId();
    const filter = encodeURIComponent(`fields/ABCClaimIDLookupId eq ${claimRowId}`);
    const data = await graphFetch(
      `/sites/${siteId}/lists/${resolveListRef(INVOICE_LIST)}/items?$filter=${filter}&$expand=fields&$top=200`,
      { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
    );
    return (data.value || []).map(i => ({ id: i.id, ...i.fields }));
  } catch (e) {
    const all = await graphListAllItems(INVOICE_LIST);
    return all.filter(r => Number(r.ABCClaimIDLookupId) === Number(claimRowId));
  }
}

// Saves the claim's line items. Skipped when the claim already has rows in the
// list (so a retry can't duplicate them).
async function saveClaimInvoices(rec) {
  const claim = rec.claim;
  if (!claim || !(claim.lineItems || []).length) return;
  const claimRowId = Number(await findClaimRowId(rec));
  if (!(claimRowId > 0)) return;
  if ((await listInvoiceRows(claimRowId)).length) return;
  const cur = claimCurrency(claim);
  for (const li of claim.lineItems) {
    const fields = {
      Title: rec.refNo,
      ABCClaimIDLookupId: claimRowId,
      Description: li.description || "",
      HasReceipt: li.hasReceipt || "",
      Amount: lineAmount(li),
      Purpose: li.purpose || "",
      Other: li.attachment || "",
      TransactionType: li.transactionType || "",
      PaymentMethod: li.paymentMethod || "",
      ConversionRate: String(lineRate(li, claim))
    };
    if (li.date) fields.InvoiceDate = li.date;
    await graphCreateItem(INVOICE_LIST, fields);
  }
}

/* --------------------------------------------------------------------------
   Loading a Pre-Approval FROM SharePoint — so an approver on another computer
   (following the link in an email / Teams message, or opening the Approver
   Console) sees the request even though it was submitted in someone else's
   browser. Rebuilds the same record shape the app uses from the
   ABC Pre-Approval row + its ABC Recipient Final Expenses rows. Claims are not
   loaded yet (their line items are not saved to SharePoint).
   -------------------------------------------------------------------------- */
const SP_STATUS_TO_STAGE = { Pending: "pending", Approved: "approved", Rejected: "rejected" };
const SP_GATE0_TO_STATUS = { Pending: "Pending", Closed: "Approved", Rejected: "Rejected", Cancelled: "Cancelled" };

async function buildPreApprovalFromRow(item) {
  const f = item.fields || {};
  const hm = !f.Gate0DH; // Higher Management requests have no DH stage
  const approvals = preApprovalChainDef(hm).map(def => {
    const key = "Gate0" + gateColumnKey(def.title);
    return {
      title: def.title, role: def.role, name: f[key] || def.fixedName || "", applicable: true,
      email: f[key + "Email"] || "", position: f[key + "Position"] || def.role,
      status: SP_STATUS_TO_STAGE[f[key + "ApprovalStatus"]] || "waiting",
      date: f[key + "ApprovalDate"] ? localISODate(new Date(f[key + "ApprovalDate"])) : null,
      comments: f[key + "Comments"] || "", rejectReason: f[key + "RejectReason"] || ""
    };
  });
  const rows = await listRecipientExpenseRows(Number(item.id));
  const dir = (typeof App !== "undefined" && App.state && App.state.recipients) || [];
  const recipients = rows.map(r => {
    const d = dir.find(x => String(x.id) === String(r.RecipientIDLookupId)) || {};
    return { name: r.Title || d.name || "", position: d.position || "", company: d.company || "", relationship: d.relationship || "", isOfficial: d.isOfficial || "No", recipientId: r.RecipientIDLookupId || "" };
  });
  const who = item.createdBy && item.createdBy.user || {};
  return {
    id: "sp-" + item.id, live: true, spId: item.id, refNo: f.FCPANo,
    submittedBy: who.displayName || f.NameofRequestor || "", submitterEmail: who.email || "",
    requestor: { name: f.NameofRequestor || "", employeeNo: f.EmployeeNo || "", department: f.Department || "", position: f.Position || "", email: f.Email || "", higherManagement: hm },
    recipients, transactionTypes: Array.isArray(f.ProposedTransaction) ? f.ProposedTransaction : [],
    description: f.ProposedTransactionDescription || "", currency: f.Currency || "",
    amounts: { gifts: f.Gifts || 0, meals: f.Meals || 0, entertainment: f.Entertainment || 0, airfare: f.Airfare || 0, transportation: f.Transportation || 0, hotel: f.Hotel || 0, othersLabel: f.OthersType || "", othersAmount: f.OthersAmount || 0 },
    paymentTo: f.PaymentInfo || "", remarks: f.Remarks || "", departmentHead: f.Gate0DH || "",
    dateSubmitted: item.createdDateTime ? localISODate(new Date(item.createdDateTime)) : null,
    approvals, status: SP_GATE0_TO_STATUS[f.Gate0] || "Pending", claim: null
  };
}

// Rebuilds the submitted claim (header, approval stages, line items, recipient
// register) from the ABC Claim Form row + its ABC Invoice and ABC Recipient
// Final Expenses rows. Returns null when the claim hasn't been submitted yet
// (a claim row that is still just "Open" has nothing to review).
async function buildClaimFromRow(rec, item, recipientRows) {
  const f = item.fields || {};
  if (!f.SubmissionDate) return null;
  const hm = !!(rec.requestor && rec.requestor.higherManagement);
  const full = claimChainDef(hm, 1e12).map(def => ({ title: def.title }));
  const approvals = [];
  claimChainDef(hm, 1e12).forEach((def, i) => {
    const c = claimStageColumns(full, i);
    if (!f[c.name]) return; // this claim's amount didn't need that stage
    approvals.push({
      title: def.title, role: def.role, name: f[c.name], applicable: true,
      email: f[c.email] || "", position: f[c.position] || def.role,
      status: SP_STATUS_TO_STAGE[f[c.status]] || "waiting",
      date: f[c.date] ? localISODate(new Date(f[c.date])) : null,
      comments: f[c.comments] || "", rejectReason: f[c.reject] || ""
    });
  });
  const lines = (await listInvoiceRows(Number(item.id))).map(r => ({
    id: "sp-" + r.id, date: r.InvoiceDate ? localISODate(new Date(r.InvoiceDate)) : "",
    description: r.Description || "", hasReceipt: r.HasReceipt || "", transactionType: r.TransactionType || "",
    paymentMethod: r.PaymentMethod || "", amount: Number(r.Amount) || 0, rate: Number(String(r.ConversionRate || "").trim()) || 0,
    purpose: r.Purpose || "", attachment: r.Other || ""
  }));
  const dir = (typeof App !== "undefined" && App.state && App.state.recipients) || [];
  return {
    spId: item.id, company: f.Company || "", costBearBy: f.CostBearCompany || "",
    date: f.SubmissionDate ? localISODate(new Date(f.SubmissionDate)) : "",
    conversionRate: f.ConversionRate || "", currency: rec.currency || "",
    purposeOfClaim: f.Purposeofclaim || "", lineItems: lines, attachments: [],
    register: {
      month: f.Month || "", year: f.Year || null,
      recipients: (recipientRows || []).map(r => {
        const d = dir.find(x => String(x.id) === String(r.RecipientIDLookupId)) || {};
        return { date: r.ClaimRegisterDate ? localISODate(new Date(r.ClaimRegisterDate)) : "", name: r.Title || d.name || "", company: d.company || "", others: 0, isOfficial: d.isOfficial || "No" };
      })
    },
    approvals,
    status: { Closed: "Closed", Rejected: "Rejected", Cancelled: "Cancelled" }[f.Gate1] || "Open",
    submittedDate: f.SubmissionDate ? localISODate(new Date(f.SubmissionDate)) : null
  };
}

async function attachClaimFromSharePoint(rec) {
  const siteId = await getSiteId();
  const filter = encodeURIComponent(`fields/Title eq '${String(rec.refNo).replace(/'/g, "''")}'`);
  const data = await graphFetch(
    `/sites/${siteId}/lists/${resolveListRef("ABC Claim Form")}/items?$filter=${filter}&$expand=fields`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  const item = data.value && data.value[0];
  if (!item) return;
  rec.claimSpId = item.id;
  const rows = await listRecipientExpenseRows(Number(rec.spId));
  rec.claim = await buildClaimFromRow(rec, item, rows);
}

// Brings one request in from SharePoint (by ABC reference number) and returns it,
// or null if there is no such row. A copy already in this browser keeps its own
// data; only its stage statuses / overall status are refreshed from SharePoint.
async function importPreApprovalByRef(refNo) {
  const siteId = await getSiteId();
  const filter = encodeURIComponent(`fields/FCPANo eq '${String(refNo).replace(/'/g, "''")}'`);
  const data = await graphFetch(
    `/sites/${siteId}/lists/${resolveListRef("ABC Pre-Approval")}/items?$filter=${filter}&$expand=fields`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  const item = data.value && data.value[0];
  if (!item) return null;
  const fresh = await buildPreApprovalFromRow(item);
  await attachClaimFromSharePoint(fresh);
  return mergeImportedPreApproval(fresh);
}

function mergeImportedPreApproval(fresh) {
  const mine = Store.getByRef(fresh.refNo);
  if (!mine) { Store.state.preApprovals.unshift(fresh); Store.save(); return fresh; }
  mine.approvals = fresh.approvals; mine.status = fresh.status; mine.spId = mine.spId || fresh.spId;
  if (fresh.claimSpId) mine.claimSpId = fresh.claimSpId;
  if (fresh.claim) {
    if (!mine.claim) mine.claim = fresh.claim;
    else {
      mine.claim.approvals = fresh.claim.approvals; mine.claim.status = fresh.claim.status; mine.claim.spId = fresh.claim.spId;
      if (!(mine.claim.lineItems || []).length) mine.claim.lineItems = fresh.claim.lineItems;
    }
  }
  Store.save();
  return mine;
}

// Imports every claim that is still Pending in SharePoint (with its request).
async function importPendingClaims() {
  const siteId = await getSiteId();
  const filter = encodeURIComponent("fields/Gate1 eq 'Pending'");
  const data = await graphFetch(
    `/sites/${siteId}/lists/${resolveListRef("ABC Claim Form")}/items?$filter=${filter}&$expand=fields($select=Title)&$top=100`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  let n = 0;
  for (const item of (data.value || [])) { if (item.fields && item.fields.Title && await importPreApprovalByRef(item.fields.Title)) n++; }
  return n;
}

// Imports every Pre-Approval that is still Pending in SharePoint — what the
// Approver Console lists for people (and admins) who did not submit them.
async function importPendingPreApprovals() {
  const siteId = await getSiteId();
  const filter = encodeURIComponent("fields/Gate0 eq 'Pending'");
  const data = await graphFetch(
    `/sites/${siteId}/lists/${resolveListRef("ABC Pre-Approval")}/items?$filter=${filter}&$expand=fields&$top=100`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  let n = 0;
  for (const item of (data.value || [])) { mergeImportedPreApproval(await buildPreApprovalFromRow(item)); n++; }
  return n;
}

/* --------------------------------------------------------------------------
   Cancellation (same behaviour as the Power Automate flows: Gate0 / Gate1 =
   "Cancelled" and IsCancelled = true on the row, and every approval step
   re-checks IsCancelled before it acts).
   -------------------------------------------------------------------------- */

// Writes a cancellation to SharePoint: the Pre-Approval row (Gate 0 =
// Cancelled), and its claim row (Gate 1 = Cancelled) unless the claim was
// already Closed. Rows that were never saved are skipped — there is nothing to
// cancel there.
// The cancel reason typed in the dialog goes into the "CancelReason" column of
// each list (single line of text, max 255 characters, so it is trimmed).
// Left out when blank. SharePoint's own Modified date records when.
function cancelReasonFields(reason) {
  const r = String(reason || "").replace(/\s+/g, " ").trim().slice(0, 255);
  return r ? { CancelReason: r } : {};
}

async function saveCancellation(rec, claimWasClosed) {
  const preId = await findPreApprovalRowId(rec);
  if (preId) await graphUpdateItem("ABC Pre-Approval", preId, { Gate0: "Cancelled", IsCancelled: true, ...cancelReasonFields(rec.cancelReason) });
  if (!claimWasClosed) {
    const claimId = await findClaimRowId(rec);
    if (claimId) await graphUpdateItem("ABC Claim Form", claimId, { Gate1: "Cancelled", IsCancelled: true, ...cancelReasonFields(rec.cancelReason) });
  }
}

// Cancelling just a claim (the "Cancel" button on the Claim Form tab): only the
// claim row's Gate 1 becomes Cancelled — the Pre-Approval stays Closed, as in
// the FCPA claim flows. A claim that was never saved has no row, so nothing to do.
async function saveClaimCancellation(rec) {
  const claimId = await findClaimRowId(rec);
  if (claimId) await graphUpdateItem("ABC Claim Form", claimId, { Gate1: "Cancelled", IsCancelled: true, ...cancelReasonFields(rec.claim && rec.claim.cancelReason) });
}

// Fresh read of SharePoint to see whether this request (or, for a claim
// decision, its claim) has been cancelled — by this app, or by anyone using
// the Power App / flows. Resolves "request" if the whole Pre-Approval is
// cancelled, "claim" if only the claim is, or false. Throws if SharePoint can't
// be read.
async function cancelledScopeInSharePoint(rec, which) {
  const preId = await findPreApprovalRowId(rec);
  if (preId) {
    const f = await graphGetItemFields("ABC Pre-Approval", preId, ["IsCancelled", "Gate0"]);
    if (f.IsCancelled === true || f.Gate0 === "Cancelled") return "request";
  }
  if (which === "claim") {
    const claimId = await findClaimRowId(rec);
    if (claimId) {
      const f = await graphGetItemFields("ABC Claim Form", claimId, ["IsCancelled", "Gate1"]);
      if (f.IsCancelled === true || f.Gate1 === "Cancelled") return "claim";
    }
  }
  return false;
}

const Store = {
  state: null,

  load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      this.state = raw ? JSON.parse(raw) : seedState();
    } catch (e) {
      this.state = seedState();
    }
    if (!this.state || !Array.isArray(this.state.preApprovals)) this.state = seedState();
    this.save();
    return this.state;
  },
  save() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(this.state));
  },
  resetDemo() {
    this.state = seedState();
    this.save();
  },

  all() { return this.state.preApprovals; },
  get(id) { return this.state.preApprovals.find(r => r.id === id); },
  getByRef(refNo) { return this.state.preApprovals.find(r => r.refNo === refNo); },

  /* Approver notification directory — where a real email integration would
     read addresses from. Resolution order:
       1) the live "ABC Authority" roster, admin-managed via the Admin page
          and backed by the real "ABC Authority" SharePoint list — the
          actual source of truth once populated
       2) a guessed placeholder (data.js defaultApproverEmail), same as the
          original demo behaviour, if no roster match is found
     No email backend is wired up in this build, so nothing is actually
     sent yet; App surfaces a toast at each notify point using whichever of
     the above resolves. (The old per-browser manual override that used to
     live here — "Notification settings" — was removed 2026-09-07: it was a
     localStorage-only stand-in that predated the Admin page, and having
     both meant a stale local override could silently shadow a real change
     made in Admin. The Admin page is now the single place to change who
     gets notified.) */
  getApproverEmail(name) {
    const roster = (typeof App !== "undefined" && App.state && App.state.approvers) || [];
    const row = roster.find(r => r.name === name);
    if (row && row.email) return row.email;
    return defaultApproverEmail(name);
  },

  async createPreApproval(draft) {
    // live = submitted from this app (as opposed to the seeded demo records);
    // only live requests are ever written to SharePoint on approve/reject.
    const rec = { ...draft, id: uid(), live: true, refNo: genRefNo(draft.submitterCountry), dateSubmitted: todayISO(), status: "Pending", claim: null };
    rec.approvals = buildChain(preApprovalChainDef(isHigherManagement(rec.requestor)), chainResolveMap(rec.requestor, rec.departmentHead));
    this.state.preApprovals.unshift(rec);
    this.save();
    rec.notify = pendingNotifyTarget(rec.approvals);
    this.notifyWaiting(rec, "preapproval");

    if (typeof isGraphConnected === "function" && isGraphConnected()) {
      try {
        await pushPreApprovalToSharePoint(rec);
        this.save(); // remember rec.spId
      } catch (e) {
        // Field-name mismatches are expected until the internal SharePoint
        // column names are verified against this best-guess mapping — see
        // pushPreApprovalToSharePoint. Keep the local submission working
        // either way rather than blocking the user on a sync failure.
        console.error("SharePoint push failed for", rec.refNo, e);
        rec.syncError = (e && e.message) || String(e);
        this.save();
      }
    }
    return rec;
  },

  cancelPreApproval(id, reason) {
    const rec = this.get(id);
    if (!rec) return;
    const claimWasClosed = !!(rec.claim && rec.claim.status === "Closed");
    this.markCancelled(rec, reason || "");
    this.syncDecision(rec, () => saveCancellation(rec, claimWasClosed));
  },

  /* ---- Email + Teams notifications (see notify.js) ------------------------
     Only for requests submitted from this app, and only to real addresses
     found in the ABC Authority roster / Entra directory (never a guessed one).
     Fire-and-forget: a failed notification never affects the decision. */
  _formName(which) { return which === "claim" ? "Claim Form" : "Pre-Approval"; },
  _factLines(rec) {
    const e = s => Notify.esc(s);
    const lines = [`<b>Requestor:</b> ${e(rec.requestor && rec.requestor.name)}`];
    if (rec.submittedBy && rec.submittedBy !== (rec.requestor && rec.requestor.name)) lines.push(`<b>Submitted by:</b> ${e(rec.submittedBy)}`);
    if (rec.description) lines.push(`<b>Description:</b> ${e(rec.description)}`);
    return lines;
  },
  // Tell the stage that is now Pending that it is their turn.
  notifyWaiting(rec, which) {
    if (!rec.live || typeof Notify === "undefined" || !Notify.enabled) return;
    const stages = which === "claim" ? (rec.claim && rec.claim.approvals) : rec.approvals;
    const stage = (stages || []).find(s => s.status === "pending");
    if (!stage || !stage.email) return;
    const form = this._formName(which);
    Notify.send({
      toEmail: stage.email, rec, which,
      subject: `${form} ${rec.refNo} is waiting for your approval (${stage.title})`,
      intro: `Hi ${Notify.esc(stage.name)}, a ${form} is waiting for your approval as ${Notify.esc(stage.title)}.`,
      lines: this._factLines(rec)
    });
  },
  // Tell the submitter how it ended (rejected, or fully approved).
  notifyOutcome(rec, which, stage, action) {
    if (!rec.live || typeof Notify === "undefined" || !Notify.enabled) return;
    const to = rec.submitterEmail || (rec.requestor && rec.requestor.email);
    if (!to) return;
    const form = this._formName(which);
    const rejected = action === "reject";
    Notify.send({
      toEmail: to, rec, which,
      subject: rejected ? `${form} ${rec.refNo} was rejected (${stage.title})` : `${form} ${rec.refNo} is fully approved`,
      intro: rejected
        ? `Your ${form} was rejected by ${Notify.esc(stage.name)} (${Notify.esc(stage.title)}).`
        : `Your ${form} has been fully approved.${which === "preapproval" ? " You can now submit the claim." : ""}`,
      lines: [...this._factLines(rec), ...(rejected && stage.rejectReason ? [`<b>Reason:</b> ${Notify.esc(stage.rejectReason)}`] : [])]
    });
  },
  // After an approve/reject: reject -> tell the submitter; last approval -> tell
  // the submitter; otherwise -> tell the next approver.
  notifyAfterDecision(rec, which, stageIndex, action) {
    const stages = which === "claim" ? rec.claim.approvals : rec.approvals;
    const stage = stages[stageIndex];
    const finished = which === "claim" ? rec.claim.status === "Closed" : rec.status === "Approved";
    if (action === "reject" || finished) this.notifyOutcome(rec, which, stage, action);
    else this.notifyWaiting(rec, which);
  },

  // Cancels just the claim; the Pre-Approval stays as it is (Closed).
  cancelClaim(id, reason) {
    const rec = this.get(id);
    if (!rec) return;
    this.markClaimCancelled(rec, reason || "");
    this.syncDecision(rec, () => saveClaimCancellation(rec));
  },
  markClaimCancelled(rec, reason) {
    const claim = this.ensureClaim(rec.id);
    claim.status = "Cancelled";
    claim.cancelReason = reason;
    claim.cancelDate = todayISO();
    this.save();
  },

  // Local half of a whole-request cancellation (also used when SharePoint
  // says it was cancelled elsewhere).
  markCancelled(rec, reason) {
    rec.status = "Cancelled";
    rec.cancelReason = reason;
    rec.cancelDate = todayISO();
    if (rec.claim && rec.claim.status === "Open") rec.claim.status = "Cancelled";
    this.save();
  },

  /* Called by the Approver Console before it records an approve/reject.
     Resolves true if the request is cancelled — the local copy is then marked
     cancelled and the decision must NOT be recorded (the flows do the same:
     Terminate "Cancelled"). Only requests submitted from this app are checked.
     If SharePoint can't be read it THROWS, so the caller can refuse rather than
     approve something that may already have been withdrawn. */
  async checkCancelled(id, which) {
    const rec = this.get(id);
    if (!rec || !rec.live || typeof isGraphConnected !== "function" || !isGraphConnected()) return false;
    if (rec.status === "Cancelled") return true;
    if (which === "claim" && rec.claim && rec.claim.status === "Cancelled") return true;
    const scope = await cancelledScopeInSharePoint(rec, which);
    if (scope === "request") this.markCancelled(rec, "Cancelled in SharePoint (Power App or another user).");
    else if (scope === "claim") this.markClaimCancelled(rec, "Cancelled in SharePoint (Power App or another user).");
    return !!scope;
  },

  ensureClaim(id) {
    const rec = this.get(id);
    if (!rec) return null;
    if (!rec.claim) rec.claim = newBlankClaim(rec, false);
    this.save();
    return rec.claim;
  },

  saveClaimForm(id, fields) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return;
    Object.assign(rec.claim, fields);
    this.save();
  },

  addClaimLineItem(id, item) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return;
    rec.claim.lineItems.push({ id: uid(), ...item });
    this.save();
  },
  removeClaimLineItem(id, itemId) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return;
    rec.claim.lineItems = rec.claim.lineItems.filter(li => li.id !== itemId);
    this.save();
  },

  saveClaimRegister(id, fields) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return;
    Object.assign(rec.claim.register, fields);
    this.save();
  },
  addRegisterRecipient(id, recipient) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return;
    rec.claim.register.recipients.push(recipient);
    this.save();
  },
  removeRegisterRecipient(id, idx) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return;
    rec.claim.register.recipients.splice(idx, 1);
    this.save();
  },

  submitClaim(id) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return null;
    rec.claim.approvals = buildChain(claimChainDef(isHigherManagement(rec.requestor), claimActualUSD(rec.claim)), chainResolveMap(rec.requestor, rec.departmentHead));
    rec.claim.submittedDate = todayISO();
    this.save();
    this.syncDecision(rec, async () => {
      await pushClaimToSharePoint(rec);
      await saveRecipientExpenses(rec); // each recipient's equal USD share of the claim
      await saveClaimInvoices(rec);     // the claim's line items, for approvers on other computers
      this.save();
    });
    this.notifyWaiting(rec, "claim");
    return pendingNotifyTarget(rec.claim.approvals);
  },

  /* Approver actions */
  actOnPreApproval(id, stageIndex, action, comments, rejectReason) {
    const rec = this.get(id);
    if (!rec) return null;
    advanceChain(rec.approvals, stageIndex, action, comments, rejectReason);
    rec.status = chainOverallStatus(rec.approvals, "Pending", "Approved");
    this.save();
    this.syncDecision(rec, async () => {
      await savePreApprovalDecision(rec, stageIndex);
      // Last stage approved -> Gate 0 is now Closed; open the claim (Gate 1 = Open).
      if (rec.status === "Approved") await createOpenClaimRow(rec);
      this.save();
    });
    this.notifyAfterDecision(rec, "preapproval", stageIndex, action);
    return action === "approve" ? pendingNotifyTarget(rec.approvals) : null;
  },

  /* Fire-and-forget SharePoint save for an approve/reject. The decision is
     already applied locally (so the UI never waits on the network); if the
     SharePoint write fails the request keeps working locally, the error is
     remembered on rec.syncError, and the user gets a toast. */
  syncDecision(rec, saveFn) {
    if (!rec.live || typeof isGraphConnected !== "function" || !isGraphConnected()) return;
    saveFn().then(() => {
      if (rec.syncError) { rec.syncError = ""; this.save(); }
    }).catch(e => {
      console.error("SharePoint save failed for", rec.refNo, e);
      rec.syncError = (e && e.message) || String(e);
      this.save();
      if (typeof App !== "undefined" && App.toast) App.toast("Decision recorded here, but saving it to SharePoint failed — see console.");
    });
  },
  actOnClaim(id, stageIndex, action, comments, rejectReason) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return null;
    advanceChain(rec.claim.approvals, stageIndex, action, comments, rejectReason);
    rec.claim.status = chainOverallStatus(rec.claim.approvals, "Open", "Closed");
    this.save();
    this.syncDecision(rec, () => saveClaimDecision(rec, stageIndex).then(() => this.save()));
    this.notifyAfterDecision(rec, "claim", stageIndex, action);
    return action === "approve" ? pendingNotifyTarget(rec.claim.approvals) : null;
  },

  /* Everything a given approver identity currently needs to act on. */
  approverQueue(identity) {
    const items = [];
    this.state.preApprovals.forEach(rec => {
      if (rec.status === "Pending") {
        rec.approvals.forEach((s, idx) => {
          if (s.status === "pending" && s.name === identity.name) {
            items.push({ flow: "preapproval", rec, stageIndex: idx, stage: s });
          }
        });
      }
      if (rec.claim && rec.claim.status === "Open" && rec.claim.approvals && rec.claim.approvals.length) {
        rec.claim.approvals.forEach((s, idx) => {
          if (s.status === "pending" && s.name === identity.name) {
            items.push({ flow: "claim", rec, stageIndex: idx, stage: s });
          }
        });
      }
    });
    return items;
  },

  /* Every stage across every record currently pending on ANYONE — the admin
     oversight equivalent of approverQueue(), used when the signed-in user is
     in ADMIN_USERS (data.js) rather than filtered to a single approver's own
     name. Read-only: seeing an item here never grants the ability to act on
     it — App._actingStage() in app.js still requires the real named approver
     to be signed in before the Approve/Reject panel appears. */
  allPendingApprovals() {
    const items = [];
    this.state.preApprovals.forEach(rec => {
      if (rec.status === "Pending") {
        rec.approvals.forEach((s, idx) => {
          if (s.status === "pending") items.push({ flow: "preapproval", rec, stageIndex: idx, stage: s });
        });
      }
      if (rec.claim && rec.claim.status === "Open" && rec.claim.approvals && rec.claim.approvals.length) {
        rec.claim.approvals.forEach((s, idx) => {
          if (s.status === "pending") items.push({ flow: "claim", rec, stageIndex: idx, stage: s });
        });
      }
    });
    return items;
  },

  /* Stages a given approver identity has already decided (approved or
     rejected) — their personal "passed" history, most recent decision first. */
  approverHistory(identity) {
    const items = [];
    this.state.preApprovals.forEach(rec => {
      rec.approvals.forEach((s, idx) => {
        if ((s.status === "approved" || s.status === "rejected") && s.name === identity.name) {
          items.push({ flow: "preapproval", rec, stageIndex: idx, stage: s });
        }
      });
      if (rec.claim && rec.claim.approvals) {
        rec.claim.approvals.forEach((s, idx) => {
          if ((s.status === "approved" || s.status === "rejected") && s.name === identity.name) {
            items.push({ flow: "claim", rec, stageIndex: idx, stage: s });
          }
        });
      }
    });
    items.sort((a, b) => (a.stage.date < b.stage.date ? 1 : -1));
    return items;
  },

  /* Every already-decided stage across every record and every approver — the
     admin oversight equivalent of approverHistory(). */
  allApprovalHistory() {
    const items = [];
    this.state.preApprovals.forEach(rec => {
      rec.approvals.forEach((s, idx) => {
        if (s.status === "approved" || s.status === "rejected") items.push({ flow: "preapproval", rec, stageIndex: idx, stage: s });
      });
      if (rec.claim && rec.claim.approvals) {
        rec.claim.approvals.forEach((s, idx) => {
          if (s.status === "approved" || s.status === "rejected") items.push({ flow: "claim", rec, stageIndex: idx, stage: s });
        });
      }
    });
    items.sort((a, b) => (a.stage.date < b.stage.date ? 1 : -1));
    return items;
  }
};
