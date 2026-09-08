/* ==========================================================================
   ABC Application — data store
   localStorage-backed model of pre-approvals & claims, standing in for the
   SharePoint list / Power Automate backend described in the release plan.
   ========================================================================== */

const STORAGE_KEY = "abc_app_state_v2";

function uid() {
  return Math.random().toString(36).slice(2, 10);
}
function genRefNo() {
  return "ABC-MY-" + uid();
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

/* Deterministic pseudo-random generator seeded from a string, so each
   recipient's mock 6-month history stays stable across re-renders. */
function seededRand(seedStr) {
  let h = 1779033703 ^ seedStr.length;
  for (let i = 0; i < seedStr.length; i++) {
    h = Math.imul(h ^ seedStr.charCodeAt(i), 3432918353);
    h = (h << 13) | (h >>> 19);
  }
  return function () {
    h = Math.imul(h ^ (h >>> 16), 2246822507);
    h = Math.imul(h ^ (h >>> 13), 3266489909);
    h ^= h >>> 16;
    return (h >>> 0) / 4294967296;
  };
}

function sixMonthHistory(recipient) {
  const rand = seededRand(recipient.company + "|" + recipient.name);
  const rows = [];
  const count = 3 + Math.floor(rand() * 3);
  for (let i = 0; i < count; i++) {
    const monthsAgo = Math.floor(rand() * 6);
    const d = new Date();
    d.setMonth(d.getMonth() - monthsAgo);
    d.setDate(1 + Math.floor(rand() * 27));
    const iso = d.toISOString().slice(0, 10);
    const pick = () => (rand() < 0.45 ? 0 : Math.round(rand() * 1400 * 100) / 100);
    const row = {
      date: iso,
      currency: rand() < 0.7 ? "USD" : "MYR",
      gifts: pick(), meals: pick(), travel: pick(), entertainment: pick(), others: pick()
    };
    rows.push(row);
  }
  rows.sort((a, b) => (a.date < b.date ? 1 : -1));
  return rows;
}

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
    rejectReason: ""
  }));
  if (stages.length) stages[0].status = "pending";
  return stages;
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
    refNo: genRefNo(),
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
    company: prefill ? "RGB" : "",
    costBearBy: prefill ? "RGB" : "",
    date: prefill ? preApproval.dateSubmitted : "",
    conversionRate: prefill ? 4.2 : "",
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
      amountMYR: 620, rate: 4.2, purpose: "Client engagement", attachment: "receipt-0" + (i + 1) + ".pdf"
    }));
    claim.register.month = "June";
    claim.approvals = buildChain(claimChainDef(isHigherManagement(preApproval.requestor), claimActualUSD(claim)), chainResolveMap(preApproval.requestor, preApproval.departmentHead));
  }
  return claim;
}

function claimActualMYR(claim) {
  return claim.lineItems.reduce((s, li) => s + (Number(li.amountMYR) || 0), 0);
}
function claimActualUSD(claim) {
  const myr = claimActualMYR(claim);
  const rate = Number(claim.conversionRate) || 0;
  return rate > 0 ? myr / rate : 0;
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
    recipients: [{ name: "Ms Gladys Lei", company: "IGT", position: "Regional Manager", relationship: "Partner", isOfficial: "No" }],
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

function gateFields(gatePrefix, stage) {
  const key = gatePrefix + gateColumnKey(stage.title);
  return {
    [key]: stage.name || "",
    [key + "Email"]: stage.email || "",
    [key + "Position"]: stage.position || "",
    [key + "ApprovalStatus"]: stage.status,
    [key + "ApprovalDate"]: stage.date,
    [key + "Comments"]: stage.comments || "",
    [key + "RejectReason"]: stage.rejectReason || ""
  };
}

async function pushPreApprovalToSharePoint(rec) {
  const a = rec.amounts;
  const fields = {
    FCPANo: rec.refNo,
    NameofRequestor: rec.requestor.name,
    SubmittedBy: rec.submittedBy || rec.requestor.name, // who actually filled in and submitted the form — may differ from the requestor (e.g. an assistant submitting on behalf of the COO)
    EmployeeNo: rec.requestor.employeeNo,
    Department: rec.requestor.department,
    Position: rec.requestor.position,
    Email: rec.requestor.email || "",
    ProposedTransaction: rec.transactionTypes || [], // MultiChoice field — array, not a joined string
    ProposedTransactionDescription: rec.description || "",
    Currency: rec.currency,
    Gifts: a.gifts, Meals: a.meals, Entertainment: a.entertainment, Airfare: a.airfare,
    Transportation: a.transportation, Hotel: a.hotel,
    OthersType: a.othersLabel, OthersAmount: a.othersAmount,
    PaymentInfo: rec.paymentTo, Remarks: rec.remarks,
    Gate0: rec.status,
    IsCancelled: false
  };
  rec.approvals.forEach(stage => Object.assign(fields, gateFields("Gate0", stage)));

  const item = await graphCreateItem("ABC Pre-Approval", fields);
  rec.spId = item.id;

  // Recipient rows: recipients are now picked from the live "FCPA Customer"
  // SharePoint list (see graph.js/app.js), and each carries the real list
  // item id through as r.recipientId when picked that way. Still written
  // without a RecipientID field below, though, because the real column
  // name on "ABC Recipient Final Expenses" itself isn't confirmed yet —
  // add it here (e.g. RecipientID: r.recipientId) once that list's schema
  // is confirmed via Graph Explorer, the same way FCPA Customer's was.
  for (const r of rec.recipients) {
    await graphCreateItem("ABC Recipient Final Expenses", {
      FCPANo: rec.refNo,
      Gifts: 0, Meals: 0, Entertainment: 0, Travel: 0, Others: 0
    });
  }
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
    const rec = { ...draft, id: uid(), refNo: genRefNo(), dateSubmitted: todayISO(), status: "Pending", claim: null };
    rec.approvals = buildChain(preApprovalChainDef(isHigherManagement(rec.requestor)), chainResolveMap(rec.requestor, rec.departmentHead));
    this.state.preApprovals.unshift(rec);
    this.save();
    rec.notify = pendingNotifyTarget(rec.approvals);

    if (typeof isGraphConnected === "function" && isGraphConnected()) {
      try {
        await pushPreApprovalToSharePoint(rec);
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
    rec.status = "Cancelled";
    rec.cancelReason = reason || "";
    rec.cancelDate = todayISO();
    if (rec.claim && rec.claim.status === "Open") rec.claim.status = "Cancelled";
    this.save();
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
    return pendingNotifyTarget(rec.claim.approvals);
  },

  /* Approver actions */
  actOnPreApproval(id, stageIndex, action, comments, rejectReason) {
    const rec = this.get(id);
    if (!rec) return null;
    advanceChain(rec.approvals, stageIndex, action, comments, rejectReason);
    rec.status = chainOverallStatus(rec.approvals, "Pending", "Approved");
    this.save();
    return action === "approve" ? pendingNotifyTarget(rec.approvals) : null;
  },
  actOnClaim(id, stageIndex, action, comments, rejectReason) {
    const rec = this.get(id);
    if (!rec || !rec.claim) return null;
    advanceChain(rec.claim.approvals, stageIndex, action, comments, rejectReason);
    rec.claim.status = chainOverallStatus(rec.claim.approvals, "Open", "Closed");
    this.save();
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
  }
};
