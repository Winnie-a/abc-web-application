/* ==========================================================================
   ABC Application — mock master data
   Standing in for the corporate directory / SSO / recipient master that a
   real deployment would pull from HR & finance systems.

   Approval hierarchy & thresholds below reflect "ABC Automation MVP —
   Release Planning v2.2" (Phase 2): updated DH routing for SSM/TSM,
   Compliance Committee 1 & 2 added to the pre-approval flow, and MD/EXCO
   Members conditionally added to the claim flow above GCOO based on the
   claim total — see claimChainDef() below for the confirmed thresholds.
   ========================================================================== */

const CURRENT_USER = {
  employeeNo: "RGB-1042",
  name: "Winnie Ang",
  department: "ENGINEERING - SYSTEM",
  position: "System Support Engineer",
  email: "winnieang@rgbgames.com",
  team: "SSM",
  higherManagement: false,
  country: "MY" // Entra ID "Country" profile field — used to tag demo/seed ABC ref numbers (see genRefNo() in store.js); real sign-in reads this live from Graph instead.
};

const EMPLOYEES = [
  CURRENT_USER,
  { employeeNo: "RGB-2031", name: "Kenix Yong", department: "FINANCE", position: "Finance Executive", email: "kenix.yong@rgbgames.com", team: "TSM", higherManagement: false },
  { employeeNo: "RGB-3110", name: "Chuah Siong Lim", department: "SALES - INTERNATIONAL", position: "Business Development Manager", email: "sl.chuah@rgbgames.com", team: "TSM", higherManagement: false },
  { employeeNo: "RGB-4402", name: "Kate Kee", department: "MARKETING", position: "Marketing Executive", email: "kate.kee@rgbgames.com", team: "SSM", higherManagement: false },
  /* "Higher Management" requestor — per the approval-flow charts, these
     employees skip the line-manager stages (DH, or HOD/SHOD/DH) and their
     requests start straight at CFO. */
  { employeeNo: "RGB-0500", name: "Ong Siew Foong", department: "MARKETING - MANAGEMENT", position: "VP, Marketing", email: "ong.foong@rgbgames.com", team: null, higherManagement: true }
];

/* People allowed to use "Can't find the recipient? Click here" on Tab B
   (Recipient(s) Info) to add an ad hoc recipient that isn't in the live
   FCPA Customer directory. Everyone else must pick an existing recipient
   from the drop-down only, so the recipient master data (and its
   Relationship with RGB / Official flags — no longer end-user-editable,
   see App.renderTabB) stays consistent with what compliance has on file.
   Match on email (lowercase) since that's stable across the static
   EMPLOYEES list and the live Graph directory. TODO(Winnie): confirm the
   real names/emails that should be on this list — currently just a
   placeholder. */
const RECIPIENT_OVERRIDE_ALLOWLIST = [
  CURRENT_USER.email
];
function isRecipientOverrideAllowed(employee) {
  if (!employee || !employee.email) return false;
  const email = employee.email.toLowerCase();
  return RECIPIENT_OVERRIDE_ALLOWLIST.some(e => e.toLowerCase() === email);
}

function employeeByNo(employeeNo) {
  return EMPLOYEES.find(e => e.employeeNo === employeeNo) || null;
}
function isHigherManagement(requestor) {
  if (!requestor) return false;
  // A requestor built from a real Graph directory pick already carries its
  // own higherManagement flag (from the FCPA Higher Management Entra ID
  // group) — trust it. Falls back to the static EMPLOYEES lookup only if
  // that flag is somehow missing (there's no unauthenticated demo/preview
  // sign-in anymore — every session is a real Microsoft account).
  if (typeof requestor.higherManagement === "boolean") return requestor.higherManagement;
  const emp = employeeByNo(requestor.employeeNo);
  return !!(emp && emp.higherManagement);
}

/* Recipient directory — company + name + position, as seen in the recipient
   drop-down and the 6-month record. This is now only a FALLBACK: the app
   pulls the real recipient list live from the "FCPA Customer" SharePoint
   list on sign-in (see graphGetRecipientDirectory() in graph.js and
   App.recipientDirectory() in app.js) and only falls back to this static
   array if that fetch comes back empty. */
const RECIPIENTS = [
  { company: "Penang", name: "Chow Bong Beng", position: "Guest" },
  { company: "Denver", name: "Chee KeXin", position: "Guest" },
  { company: "Tencent", name: "Tendu France", position: "Partner Representative" },
  { company: "Zai3Bei", name: "Zai 3 Bei", position: "Guest" },
  { company: "PalmGold Corporate Services Sdn Bhd (PG2)", name: "Bryan Ng Chi Che", position: "Director" },
  { company: "PalmGold Corporate Services Sdn Bhd (PG2)", name: "Mary Yeap Mei Pheng", position: "Finance Manager" },
  { company: "PalmGold Corporate Services Sdn Bhd (PG2)", name: "Gary Lai Voon Yaw", position: "Operations Manager" },
  { company: "Light&Wonder", name: "John Preston", position: "Account Manager" },
  { company: "PalmGold Corporate Services Sdn Bhd (PG2)", name: "Bong Kim Ann", position: "Executive" },
  { company: "PalmGold Corporate Services Sdn Bhd (PG2)", name: "Andy Koh Cheng Siew", position: "Executive" },
  { company: "IGT", name: "Ms Gladys Lei", position: "Regional Manager" },
  { company: "RGB Sdn. Bhd.", name: "Kate Kee", position: "Marketing Executive" },
  { company: "RGB SDN BHD", name: "Kenix Yong", position: "Finance Executive" },
  { company: "RGB SDN BHD", name: "Chuah Siong Lim", position: "Business Development Manager" },
  { company: "IGT Solutions Pvt Ltd (Malaysia)", name: "Ankur Sharma", position: "Solutions Consultant" },
  { company: "D'heights Resort & Casino Clark Philippines", name: "Lee Seong Gi", position: "President" },
  { company: "D'heights Resort & Casino Clark Philippines", name: "Lee Sung Gyun", position: "Finance Manager" },
  { company: "D'heights Resort & Casino Clark Philippines", name: "Lee Seong Jin", position: "Property Sales Manager" },
  { company: "The Grand Ho Tram Strip Casino", name: "Charles Seo", position: "General Manager" },
  { company: "Thunderbird Resort & Casino", name: "Lourdes Sarah Cano Garcia", position: "Marketing Director" },
  { company: "Resort World Malaysia", name: "Lee Hoon Kim", position: "VIP Relations Manager" },
  { company: "StarDream Cruises", name: "Shi Qi Jian", position: "Casino Manager" },
  { company: "StarDream Cruises", name: "Chuah CK", position: "Executive" }
];

const RELATIONSHIP_OPTIONS = ["Customer", "Supplier", "Partner", "Government Official", "Other"];

const COMPANIES = ["RGB", "PalmGold Corporate Services Sdn Bhd (PG2)", "IGT", "Light & Wonder"];

/* Phase 2: expanded currency support (previously MYR, USD, VND, PHP, THB, EUR). */
const CURRENCIES = ["USD", "MYR", "VND", "PHP", "THB", "EUR", "SGD", "GBP", "MOP", "AUD", "HKD", "KRW", "LKR"];

/* Approximate, static FX-to-USD table — for MVP demo threshold/summary math only. */
const FX_TO_USD = {
  USD: 1, MYR: 0.22, VND: 0.00004, PHP: 0.017, THB: 0.028, EUR: 1.09,
  SGD: 0.75, GBP: 1.27, MOP: 0.125, AUD: 0.66, HKD: 0.128, KRW: 0.00072, LKR: 0.0034
};
function toUSD(amount, currency) {
  const rate = FX_TO_USD[currency];
  return rate ? amount * rate : amount;
}

const TRANSACTION_TYPES = ["Gifts", "Meals", "Travel", "Entertainment", "Others"];

/* Tab D transaction type -> which Tab E amount field(s) it unlocks */
const TYPE_TO_AMOUNT_FIELDS = {
  Gifts: ["gifts"],
  Meals: ["meals"],
  Travel: ["airfare", "transportation", "hotel"],
  Entertainment: ["entertainment"],
  Others: ["othersLabel", "othersAmount"]
};

const AMOUNT_FIELDS = [
  { key: "gifts", label: "Gifts" },
  { key: "meals", label: "Meals" },
  { key: "entertainment", label: "Entertainment" },
  { key: "airfare", label: "Airfare" },
  { key: "transportation", label: "Transportation" },
  { key: "hotel", label: "Hotel" }
];

const PAYMENT_METHODS = ["Cash", "Bank Transfer", "Credit Card", "Cheque"];

/* Phase 2: updated Department Head selection list — Chow Bong Weng now routes
   SSM requests, Lim Cheng Soon routes TSM requests; the requestor's team
   (see EMPLOYEES) is used to pre-select the likely DH, but the field stays a
   manual, overridable drop-down as in the original UI. */
const DEPARTMENT_HEADS = [
  { name: "Chow Bong Weng", team: "SSM" },
  { name: "Lim Cheng Soon", team: "TSM" },
  { name: "Chuah Eng Hwa" },
  { name: "Dato' Chuah Kim Chiew" },
  { name: "Khaw Chai Huat" }
];

/* Claim-flow reporting line below the Department Head — Head of Department
   (HOD) then Senior Head of Department (SHOD) — per team, feeding the
   "Claim Form Approval Flow" chart's Standard Employees branch. Requests
   from a team with no specific line use the DEFAULT line. */
const ORG_LINE = {
  SSM: { hod: "Tan Wei Loon", shod: "Ong Bee Leng" },
  TSM: { hod: "Farah Aziz", shod: "Vince Tan" },
  DEFAULT: { hod: "Chuah Eng Hwa", shod: "Dato' Chuah Kim Chiew" }
};
function orgLineFor(team) {
  return ORG_LINE[team] || ORG_LINE.DEFAULT;
}

/* --------------------------------------------------------------------------
   Approval chains — per the "Pre-Approval Form Approval Flow" and "Claim
   Form Approval Flow" charts. Both branch on whether the requestor is
   Higher Management (see isHigherManagement): Higher Management requests
   skip the line-manager stage(s) and start straight at CFO.

     Pre-Approval · Standard:   DH -> CFO -> GCOO -> Compliance Committee 1 -> Compliance Committee 2
     Pre-Approval · Higher Mgmt:      CFO -> GCOO -> Compliance Committee 1 -> Compliance Committee 2
   Every Pre-Approval stage is required, in order — no amount-based gating.

     Claim · Standard:  HOD -> SHOD -> DH -> CFO -> GCOO -> [MD] -> [EXCO Members]
     Claim · Higher Mgmt:            CFO -> GCOO -> [MD] -> [EXCO Members]
   The Claim flow DOES have amount-based gating (confirmed with Winnie
   2026-09-07, against the total claim amount converted to USD via
   claimActualUSD()): everyone always goes up to GCOO; MD and then EXCO
   Members are conditionally appended on top —
     <= USD 7,500              -> stop after GCOO
     USD 7,501 - 25,000        -> also requires MD
     USD 25,001 and above      -> also requires EXCO Members, where EVERY
                                   EXCO roster member must individually
                                   approve (not one collective sign-off) —
                                   built below as one sequential stage per
                                   member via excoStageDefs().
   -------------------------------------------------------------------------- */
const CLAIM_TIER_MD_ABOVE_USD = 7500;    // above this, MD approval is also required
const CLAIM_TIER_EXCO_ABOVE_USD = 25000; // above this, every EXCO member must also approve

// Approval order per Winnie 2026-09-08: EXCO Members sign off from the most
// junior seat up to the most senior (SVP first .. MD last), the reverse of
// the roster's numbered list (MD/GCOO/COO/ED/ED/CFO/SVP).
const EXCO_ROSTER = [
  { name: "Mazlan Bin Ismail", title: "SVP" },
  { name: "Liew Yung Kuan", title: "CFO" },
  { name: "Ganaser A/L Kaliappen", title: "ED" },
  { name: "Chuah Hui Jing", title: "ED" },
  { name: "Chuah Eng Meng", title: "COO" },
  { name: "Datuk Steven Lim", title: "GCOO" },
  { name: "Dato' Seri Chuah Kim Seah", title: "MD" }
];

const STAGE_HOD = { title: "HOD", role: "Head of Department", resolve: "hod" };
const STAGE_SHOD = { title: "SHOD", role: "Senior Head of Department", resolve: "shod" };
const STAGE_DH = { title: "DH", role: "Department Head", resolve: "dh" };
const STAGE_CFO = { title: "CFO", role: "Chief Financial Officer", fixedName: "Liew Yung Kuan" };
const STAGE_GCOO = { title: "GCOO", role: "Group Chief Operating Officer", fixedName: "Datuk Steven Lim" };
const STAGE_MD = { title: "MD", role: "Managing Director", fixedName: "Dato' Seri Chuah Kim Seah" };
const STAGE_COMPLIANCE1 = { title: "Compliance Committee 1", role: "Compliance Committee 1", fixedName: "Nicole Chan" };
const STAGE_COMPLIANCE2 = { title: "Compliance Committee 2", role: "Compliance Committee 2", fixedName: "Ganaser A/L Kaliappen" };

/* One sequential approval stage per EXCO roster member (title "EXCO - MD",
   "EXCO - GCOO", ...), so every member gets their own row in the Approver
   Console and must individually approve — instead of a single collective
   "EXCO Members" placeholder stage. Takes an optional live roster (shape:
   {name, title|role}) for when the "ABC Authority" SharePoint list's real
   schema is confirmed and can reliably tell EXCO seats apart from
   Compliance/other authority roles; falls back to the static EXCO_ROSTER
   (used today, since that isn't wired up yet). */
function excoStageDefs(authorityList) {
  const roster = (authorityList && authorityList.length)
    ? authorityList.map(a => ({ name: a.name, title: a.tag || a.title || a.role || a.position || "EXCO" }))
    : EXCO_ROSTER.map(e => ({ name: e.name, title: e.title }));
  return roster.map(m => ({ title: "EXCO - " + m.title, role: "EXCO Members", fixedName: m.name }));
}

function preApprovalChainDef(higherMgmt) {
  return [
    ...(higherMgmt ? [] : [STAGE_DH]),
    STAGE_CFO, STAGE_GCOO, STAGE_COMPLIANCE1, STAGE_COMPLIANCE2
  ];
}
function claimChainDef(higherMgmt, totalUSD) {
  const chain = [
    ...(higherMgmt ? [] : [STAGE_HOD, STAGE_SHOD, STAGE_DH]),
    STAGE_CFO, STAGE_GCOO
  ];
  const amt = Number(totalUSD) || 0;
  if (amt > CLAIM_TIER_MD_ABOVE_USD) chain.push(STAGE_MD);
  if (amt > CLAIM_TIER_EXCO_ABOVE_USD) chain.push(...excoStageDefs());
  return chain;
}

/* All possible approver identities an "Approver" can sign in as: every DH,
   every HOD/SHOD across the org lines, plus every fixed-name role. */
function allApproverIdentities() {
  const seen = new Set();
  const list = [];
  function add(name, title, role) {
    if (name && !seen.has(name)) { seen.add(name); list.push({ name, title, role }); }
  }
  DEPARTMENT_HEADS.forEach(dh => add(dh.name, "DH", "Department Head"));
  Object.values(ORG_LINE).forEach(line => {
    add(line.hod, "HOD", "Head of Department");
    add(line.shod, "SHOD", "Senior Head of Department");
  });
  [STAGE_CFO, STAGE_GCOO, STAGE_MD, STAGE_COMPLIANCE1, STAGE_COMPLIANCE2].forEach(s => add(s.fixedName, s.title, s.role));
  // Every individual EXCO roster member also needs to be able to sign in as
  // themselves and act on their own "EXCO - <role>" stage (see
  // excoStageDefs() above) — not just the fixed CFO/GCOO/MD/Compliance roles.
  EXCO_ROSTER.forEach(m => add(m.name, "EXCO - " + m.title, "EXCO Members"));
  return list;
}

function defaultDHForEmployee(emp) {
  const match = DEPARTMENT_HEADS.find(dh => dh.team && emp && dh.team === emp.team);
  return match ? match.name : "";
}

/* --------------------------------------------------------------------------
   Approver email directory
   The final fallback for "who gets notified" once the live "ABC Authority"
   roster (see Store.getApproverEmail in store.js) has no match for a given
   approver name. This app has no backend, so nothing is actually emailed —
   these addresses are what a real send integration would read once wired
   up. Seed values are a placeholder guess at a first.last@rgbgames.com
   pattern; edit them here in data.js, or manage real addresses via the
   Admin page (backed by the "ABC Authority" SharePoint list).
   -------------------------------------------------------------------------- */
function slugEmail(name) {
  const clean = name
    .replace(/\(.*?\)/g, "")
    .replace(/^(Dato'?\s*Seri|Dato'?|Datuk|Mr\.?|Ms\.?)\s+/i, "")
    .replace(/[^a-zA-Z\s]/g, "")
    .trim()
    .split(/\s+/);
  if (clean.length < 2) return (clean[0] || "user").toLowerCase() + "@rgbgames.com";
  return (clean[0] + "." + clean[clean.length - 1]).toLowerCase() + "@rgbgames.com";
}

const APPROVER_EMAILS_SEED = {
  "Board of Directors": "board-secretary@rgbgames.com"
};

function defaultApproverEmail(name) {
  return APPROVER_EMAILS_SEED[name] || slugEmail(name);
}
