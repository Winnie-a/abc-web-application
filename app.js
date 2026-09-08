/* ==========================================================================
   ABC Application — UI layer (vanilla JS, no build step / no framework)
   ========================================================================== */

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function initials(name) {
  return (name || "?").split(" ").filter(Boolean).slice(0, 2).map(p => p[0]).join("").toUpperCase();
}
function pillClass(status) {
  const s = (status || "").toLowerCase();
  if (s === "pending" || s === "open" || s === "waiting") return "pill-pending";
  if (s === "approved" || s === "closed") return "pill-approved";
  if (s === "rejected") return "pill-rejected";
  return "pill-pending";
}
function stagePillClass(status) {
  if (status === "approved") return "pill-approved";
  if (status === "rejected") return "pill-rejected";
  if (status === "pending") return "pill-pending";
  return "";
}
function stageLabel(status) {
  return { not_required: "Not Required", waiting: "-", pending: "Pending", approved: "Approved", rejected: "Rejected" }[status] || "-";
}

const App = {
  state: {
    session: null,
    view: "login",
    wizard: null,
    sub: null,
    detail: null,
    approver: null,
    modal: null,
    authBusy: false,
    dhOptions: [], // live FCPA DH group members, populated on real sign-in — see signInWithMicrosoft()
    directory: [], // live staff directory (every user), populated on real sign-in — see signInWithMicrosoft()
    approvers: [], // live "ABC Authority" roster (CFO/GCOO/MD/EXCO/Compliance name+email) — read-only here; edited directly in the SharePoint list
    recipients: [] // live "FCPA Customer" roster (recipient name/position/company/relationship), populated on real sign-in — see signInWithMicrosoft()
  },

  init() {
    Store.load();
    this.render();
  },

  toast(msg) {
    let t = document.getElementById("toast");
    if (!t) {
      t = document.createElement("div");
      t.id = "toast";
      document.body.appendChild(t);
    }
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(this._toastTimer);
    this._toastTimer = setTimeout(() => t.classList.remove("show"), 2600);
  },

  signOut() {
    this.state.session = null;
    this.state.view = "login";
    this.state.wizard = null;
    this.state.sub = null;
    this.state.detail = null;
    this.state.approver = null;
    this.render();
  },

  goHome() {
    this.state.view = "home";
    this.state.detail = null;
    this.render();
  },

  /* Whether the signed-in Microsoft account is also one of the named
     approver roles (CFO, GCOO, MD, EXCO, Compliance Committee, DH,
     HOD/SHOD, ...) from data.js — decides whether the Approver Console nav
     link shows up. Replaces the old, unauthenticated "Approver sign-in"
     picker: acting on a stage now requires actually being signed in as
     that real person, not just selecting their name from a dropdown. */
  isApproverUser() {
    const s = this.state.session;
    if (!s || !s.employee) return false;
    return allApproverIdentities().some(i => i.name === s.employee.name);
  },
  approverIdentity() {
    const s = this.state.session;
    if (!s || !s.employee) return null;
    const match = allApproverIdentities().find(i => i.name === s.employee.name);
    return { name: s.employee.name, role: match ? match.role : (s.employee.position || "") };
  },
  openApproverConsole() {
    if (!this.isApproverUser()) { this.toast("Your signed-in account isn't in the approver roster."); return; }
    this.state.view = "approverHome";
    this.render();
  },

  async signInWithMicrosoft() {
    if (this.state.authBusy) return;
    this.state.authBusy = true;
    this.render();
    try {
      await signIn();
      const [me, directory, dhOptions, hmMembers, approvers, recipients] = await Promise.all([
        graphGetMe(),
        graphListUsers(),
        graphGetGroupMembers(GRAPH_GROUPS.FCPA_DH),
        graphGetGroupMembers(GRAPH_GROUPS.HIGHER_MANAGEMENT),
        // Don't let the "ABC Authority" list (e.g. unconfirmed column
        // names — see graph.js) block sign-in for everyone. Fall back to
        // an empty roster and let the Admin page (or the notification
        // email lookup in store.js) show that gap instead of the whole
        // app becoming unusable.
        graphGetApproverRoster().catch(e => { console.error("ABC Authority list fetch failed:", e); return []; }),
        // Same defensive fallback for "FCPA Customer" — recipientDirectory()
        // below falls back to the static RECIPIENTS list in data.js if this
        // comes back empty, so a fetch failure degrades gracefully instead
        // of blocking sign-in or emptying the recipient picker.
        graphGetRecipientDirectory().catch(e => { console.error("FCPA Customer list fetch failed:", e); return []; })
      ]);
      const hmIds = new Set(hmMembers.map(u => u.graphId));
      directory.forEach(u => { u.higherManagement = hmIds.has(u.graphId); });
      me.higherManagement = hmIds.has(me.graphId);
      this.state.directory = directory;
      this.state.dhOptions = dhOptions;
      this.state.approvers = approvers;
      this.state.recipients = recipients;
      this.state.session = { employee: me };
      this.state.view = "home";
    } catch (e) {
      this.toast("Sign-in failed: " + (e.message || e));
    } finally {
      this.state.authBusy = false;
      this.render();
    }
  },

  /* ======================= RENDER DISPATCH ======================= */

  render() {
    const app = document.getElementById("app");
    let body = "";
    if (this.state.view === "login") body = this.renderLogin();
    else if (this.state.view === "home") body = this.topbar() + this.renderHome();
    else if (this.state.view === "newRequest") body = this.topbar() + this.renderNewRequest();
    else if (this.state.view === "submission") body = this.topbar() + this.renderSubmission();
    else if (this.state.view === "detail") body = this.topbar() + this.renderDetailPage();
    else if (this.state.view === "approverHome") body = this.topbar() + this.renderApproverHome();
    app.innerHTML = body + this.renderModal() + `<footer class="appfoot">ABC Application &middot; Pre-Approval, Register &amp; Claims Submission System &middot; <a onclick="Store.resetDemo();App.render();">Reset demo data</a></footer>`;
  },

  topbar() {
    const s = this.state.session;
    if (!s) return "";
    const who = `${esc(s.employee.name)} <span class="muted">&middot; ${esc(s.employee.employeeNo)}</span>`;
    return `
    <div class="topbar">
      <div class="brand" onclick="App.goHome()" style="cursor:pointer">ABC <small>&nbsp;Pre-Approval &amp; Claims</small></div>
      <div class="user">
        ${this.isApproverUser() ? `<a onclick="App.openApproverConsole()">Approver Console</a>` : ""}
        <span>${who}</span>
        <div class="avatar">${initials(s.employee.name)}</div>
        <button class="signout" onclick="App.signOut()">Sign out</button>
      </div>
    </div>`;
  },

  /* ======================= LOGIN ======================= */

  renderLogin() {
    return `
    <div class="login-wrap">
      <div class="login-card">
        <h1>ABC</h1>
        <p class="sub">Pre-Approval, Register &amp; Claims Submission System</p>
        <p class="small muted" style="text-align:left;margin:20px 0 20px;">Sign in with your corporate Microsoft account to continue. If you're an approver (CFO, GCOO, MD, EXCO, Compliance Committee, Department Head, HOD/SHOD...), signing in also unlocks your Approver Console automatically — there's no separate approver sign-in anymore.</p>
        <button class="btn btn-primary" style="width:100%;" onclick="App.signInWithMicrosoft()" ${this.state.authBusy ? "disabled" : ""}>${this.state.authBusy ? "Signing in&hellip;" : "Sign in with Microsoft"}</button>
        <div class="hint">Uses your real Microsoft 365 sign-in and directory profile.</div>
      </div>
    </div>`;
  },

  /* ======================= HOME ======================= */

  renderHome() {
    return `
    <div class="home-wrap">
      <div class="home-title">ABC</div>
      <div class="home-actions">
        <button class="home-btn" onclick="App.startNewRequest()">Create New<br>Request</button>
        <button class="home-btn" onclick="App.openSubmission()">Check My<br>Submission</button>
      </div>
      <div class="home-sub">Gift, Meal, Travel &amp; Entertainment Pre-Approval &amp; Claims</div>
    </div>`;
  },

  /* ======================= NEW REQUEST WIZARD ======================= */

  startNewRequest() {
    const emp = this.state.session.employee;
    this.state.wizard = {
      tab: "A",
      requestor: { name: emp.name, employeeNo: emp.employeeNo, department: emp.department, position: emp.position, higherManagement: !!emp.higherManagement },
      recipients: [],
      transactionTypes: [],
      description: "",
      currency: "",
      amounts: { gifts: 0, meals: 0, entertainment: 0, airfare: 0, transportation: 0, hotel: 0, othersLabel: "", othersAmount: 0 },
      paymentTo: "",
      remarks: "",
      departmentHead: defaultDHForEmployee(emp) || ""
    };
    this.state.view = "newRequest";
    this.render();
  },

  cancelWizard() {
    this.state.wizard = null;
    this.goHome();
  },

  setWizardTab(t) {
    this.state.wizard.tab = t;
    this.render();
  },

  wizardValidateTab(t) {
    const w = this.state.wizard;
    if (t === "A" && !w.requestor.employeeNo) return "Please select a requestor.";
    if (t === "B" && w.recipients.length === 0) return "Please add at least one recipient.";
    if (t === "D" && w.transactionTypes.length === 0) return "Please select at least one transaction type.";
    if (t === "E" && !w.currency) return "Please select a currency.";
    return null;
  },

  wizardNext() {
    const order = ["A", "B", "C", "D", "E", "F"];
    const w = this.state.wizard;
    const err = this.wizardValidateTab(w.tab);
    if (err) { this.toast(err); return; }
    const idx = order.indexOf(w.tab);
    if (idx < order.length - 1) { w.tab = order[idx + 1]; this.render(); }
  },
  wizardBack() {
    const order = ["A", "B", "C", "D", "E", "F"];
    const w = this.state.wizard;
    const idx = order.indexOf(w.tab);
    if (idx > 0) { w.tab = order[idx - 1]; this.render(); }
  },

  renderNewRequest() {
    const w = this.state.wizard;
    const tabs = [
      ["A", "Requestor Info"], ["B", "Recipient(s) Info"], ["C", "6-Month Record"],
      ["D", "Transaction Details"], ["E", "Payment Amounts"], ["F", "Remarks & Submit"]
    ];
    return `
    <div class="page">
      <div class="page-header">
        <h2>Gift, Meal, Travel &amp; Entertainment Pre-Approval Form</h2>
        <button class="btn btn-secondary" onclick="App.cancelWizard()">Back</button>
      </div>
      <div class="tabs">
        ${tabs.map(([k, label]) => `<a class="tab ${w.tab === k ? "active" : ""}" onclick="App.setWizardTab('${k}')"><span class="letter">${k}</span>${label}</a>`).join("")}
      </div>
      ${this.renderWizardTab(w.tab)}
      <div class="wizard-footer">
        <button class="btn btn-secondary" ${w.tab === "A" ? "disabled" : ""} onclick="App.wizardBack()">&larr; Back</button>
        ${w.tab === "F" ? `<button class="btn btn-primary" onclick="App.openDHModal()">Submit</button>` : `<button class="btn btn-primary" onclick="App.wizardNext()">Next &rarr;</button>`}
      </div>
    </div>`;
  },

  renderWizardTab(t) {
    const w = this.state.wizard;
    if (t === "A") return this.renderTabA(w);
    if (t === "B") return this.renderTabB(w);
    if (t === "C") return this.renderTabC(w.recipients);
    if (t === "D") return this.renderTabD(w);
    if (t === "E") return this.renderTabE(w);
    if (t === "F") return this.renderTabF(w);
    return "";
  },

  renderTabA(w) {
    const me = this.state.session.employee;
    const onBehalf = w.requestor.name && w.requestor.name !== me.name;
    return `
    <div class="card">
      <div class="field-grid">
        <div class="field combo">
          <label>Name of Requestor</label>
          <input type="text" id="reqSearch" autocomplete="off" value="${esc(w.requestor.name)}"
            oninput="App.filterRequestor(this.value)" onfocus="App.filterRequestor(this.value)"
            onblur="setTimeout(()=>{const d=document.getElementById('reqComboList'); if(d) d.style.display='none';},150)">
          <div id="reqComboList" class="combo-list" style="display:none;"></div>
        </div>
        <div class="field">
          <label>Employee No <span class="req">*</span> <span class="sub-label">(auto-populated)</span></label>
          <input type="text" value="${esc(w.requestor.employeeNo)}" disabled>
        </div>
        <div class="field">
          <label>Department</label>
          <input type="text" value="${esc(w.requestor.department)}" disabled>
        </div>
        <div class="field">
          <label>Position</label>
          <input type="text" value="${esc(w.requestor.position)}" disabled>
        </div>
        <div class="field">
          <label>Submitted By <span class="sub-label">(your signed-in account)</span></label>
          <input type="text" value="${esc(me.name)}" disabled>
        </div>
      </div>
      ${onBehalf ? `<div class="banner" style="margin-top:14px;"><span>&#8505;&#65039;</span><div><b>Submitting on behalf of ${esc(w.requestor.name)}</b>Your account (${esc(me.name)}) is recorded as the submitter; ${esc(w.requestor.name)} stays the requestor on this form and is who the approval flow evaluates (e.g. for Higher Management routing).</div></div>` : ""}
    </div>`;
  },
  requestorDirectory() {
    // Live Entra ID directory, populated on sign-in; falls back to the
    // static EMPLOYEES list only if that fetch somehow came back empty.
    return this.state.directory.length ? this.state.directory : EMPLOYEES;
  },
  filterRequestor(q) {
    const list = document.getElementById("reqComboList");
    if (!list) return;
    const items = this.requestorDirectory().filter(e => e.name.toLowerCase().includes((q || "").toLowerCase()));
    list.innerHTML = items.length
      ? items.map(e => `<div onmousedown="App.pickRequestor('${esc(e.employeeNo)}')">${esc(e.name)} <span class="muted small">&middot; ${esc(e.employeeNo)} &middot; ${esc(e.department)}</span></div>`).join("")
      : `<div class="none">No matching employee</div>`;
    list.style.display = "block";
  },
  pickRequestor(empNo) {
    const emp = this.requestorDirectory().find(e => e.employeeNo === empNo);
    if (!emp) return;
    const w = this.state.wizard;
    w.requestor = { name: emp.name, employeeNo: emp.employeeNo, department: emp.department, position: emp.position, higherManagement: !!emp.higherManagement };
    if (!w.departmentHead) w.departmentHead = defaultDHForEmployee(emp) || "";
    this.render();
  },

  renderTabB(w) {
    // Relationship with RGB / Official are now DISPLAY-ONLY here — they come
    // straight from the FCPA Customer directory (or from the "Can't find the
    // recipient?" modal for an ad hoc entry) and end users can no longer
    // edit them on this screen, so the recipient master data stays
    // consistent with what compliance actually has on file.
    const rows = w.recipients.map((r, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>${esc(r.name)}</td>
        <td>${esc(r.position)}</td>
        <td>${esc(r.company)}</td>
        <td>${esc(r.relationship || "-")}</td>
        <td>${esc(r.isOfficial || "No")}</td>
        <td><button class="btn-icon" title="Remove" onclick="App.removeRecipient(${i})">&#128465;</button></td>
      </tr>`).join("");
    // "Can't find the recipient?" (ad hoc recipient entry, bypassing the
    // live directory) is restricted to a short hardcoded allowlist — see
    // canAddAdHocRecipient() in data.js — so end users can't casually add
    // untracked recipients; everyone else must pick from the directory.
    const canAddAdHoc = this.canAddAdHocRecipient();
    return `
    <div class="card">
      <div class="section-title">
        Recipient(s) Info
        ${canAddAdHoc ? `<a onclick="App.openAddRecipientModal()">Can't find the recipient? Click here</a>` : ""}
      </div>
      <div class="banner warn" style="margin-bottom:16px;">
        <span>&#9888;</span>
        <div><b>Important</b>Select the recipient from the drop-down so their previous six (6) months of records load on the next screen.</div>
      </div>
      <div class="table-wrap">
        <table class="data">
          <thead><tr><th>No</th><th>Name of Recipient</th><th>Position</th><th>Company / Organization</th><th>Relationship with RGB</th><th>Official</th><th></th></tr></thead>
          <tbody>
            ${rows || `<tr class="empty-row"><td colspan="7">No recipients added yet.</td></tr>`}
          </tbody>
        </table>
      </div>
      <div class="combo" style="margin-top:14px;">
        <input type="text" id="recipSearch" autocomplete="off" placeholder="Find recipient"
          oninput="App.filterRecipient(this.value)" onfocus="App.filterRecipient(this.value)"
          onblur="setTimeout(()=>{const d=document.getElementById('recipComboList'); if(d) d.style.display='none';},150)">
        <div id="recipComboList" class="combo-list" style="display:none;"></div>
      </div>
    </div>`;
  },
  canAddAdHocRecipient() {
    const s = this.state.session;
    return !!(s && s.employee && isRecipientOverrideAllowed(s.employee));
  },
  recipientDirectory() {
    // Live "FCPA Customer" SharePoint list, populated on sign-in — falls
    // back to the static RECIPIENTS list in data.js only if that fetch
    // somehow came back empty (mirrors requestorDirectory() above).
    return this.state.recipients.length ? this.state.recipients : RECIPIENTS;
  },
  filterRecipient(q) {
    const list = document.getElementById("recipComboList");
    if (!list) return;
    const dir = this.recipientDirectory();
    const already = new Set(this.state.wizard.recipients.map(r => r.name + "|" + r.company));
    const items = dir.filter(r => !already.has(r.name + "|" + r.company) && (r.name.toLowerCase().includes((q || "").toLowerCase()) || r.company.toLowerCase().includes((q || "").toLowerCase())));
    const noMatchHint = this.canAddAdHocRecipient()
      ? `No matching recipient &mdash; try "Can't find the recipient? Click here"`
      : `No matching recipient.`;
    list.innerHTML = items.length
      ? items.map((r, i) => `<div onmousedown="App.pickRecipient(${dir.indexOf(r)})">${esc(r.company)} | ${esc(r.name)} <span class="muted small">&middot; ${esc(r.position)}</span></div>`).join("")
      : `<div class="none">${noMatchHint}</div>`;
    list.style.display = "block";
  },
  pickRecipient(idx) {
    const r = this.recipientDirectory()[idx];
    if (!r) return;
    this.state.wizard.recipients.push({
      name: r.name, position: r.position, company: r.company,
      // FCPA Customer rows carry real relationship/official values — prefill
      // from them when present (still user-editable via the dropdowns);
      // the static RECIPIENTS fallback has neither field, so this defaults
      // the same way it always did (blank / "No") for that path.
      relationship: RELATIONSHIP_OPTIONS.includes(r.relationship) ? r.relationship : "",
      isOfficial: r.isOfficial === "Yes" ? "Yes" : "No",
      // Carries the real FCPA Customer SharePoint item id through when the
      // pick came from the live directory (null for the static RECIPIENTS
      // fallback, which has no such id) — not used yet, but this is what
      // pushPreApprovalToSharePoint()'s RecipientID linkage (store.js) will
      // read once the "ABC Recipient Final Expenses" schema is confirmed.
      recipientId: r.id || null
    });
    const s = document.getElementById("recipSearch");
    if (s) s.value = "";
    this.render();
  },
  removeRecipient(i) {
    this.state.wizard.recipients.splice(i, 1);
    this.render();
  },

  openAddRecipientModal(forRegister) {
    this.state.modal = { type: "addRecipient", forRegister: !!forRegister, form: { name: "", position: "", company: "", relationship: "", isOfficial: "No" } };
    this.render();
  },
  submitAddRecipientModal() {
    const f = this.state.modal.form;
    if (!f.name || !f.company) { this.toast("Please enter at least the recipient's name and company."); return; }
    if (this.state.modal.forRegister) {
      Store.addRegisterRecipient(this.state.detail.id, { date: todayISO(), name: f.name, company: f.company, others: 0, isOfficial: f.isOfficial || "No" });
    } else {
      this.state.wizard.recipients.push({ ...f });
    }
    this.state.modal = null;
    this.render();
  },

  renderTabC(recipients) {
    // Stashed so the "Summary" button below works from both call sites:
    // the active wizard (renderTabC(w.recipients)) and the read-only detail
    // page for an already-submitted pre-approval (renderTabC(rec.recipients)
    // in renderDetailPreApproval) — the button used to always read
    // App.state.wizard.recipients, which is null on the detail page and made
    // it silently do nothing there. See openSixMonthSummary() below.
    this._tabCRecipients = recipients;
    if (!recipients.length) {
      return `<div class="card"><p class="muted">Add a recipient in Tab B to see their spending history here.</p></div>`;
    }
    let rows = "";
    recipients.forEach(r => {
      const hist = sixMonthHistory(r);
      if (!hist.length) {
        rows += `<tr><td>${esc(r.name)}</td><td>${esc(r.position)}</td><td colspan="7" class="muted">No prior record found for this recipient.</td></tr>`;
      }
      hist.forEach(row => {
        rows += `<tr>
          <td>${esc(r.name)}</td><td>${esc(r.position)}</td><td>${fmtDate(row.date)}</td><td>${esc(row.currency)}</td>
          <td>${fmtMoney(row.gifts)}</td><td>${fmtMoney(row.meals)}</td><td>${fmtMoney(row.travel)}</td>
          <td>${fmtMoney(row.entertainment)}</td><td>${fmtMoney(row.others)}</td><td><b>${fmtMoney(rowTotal(row))}</b></td>
        </tr>`;
      });
    });
    return `
    <div class="card">
      <div class="section-title">Record of Previous Six (6) Months of Recipient
        <button class="btn btn-secondary btn-sm" onclick="App.openSixMonthSummary()">Summary</button>
      </div>
      <div class="table-wrap">
        <table class="data">
          <thead><tr><th>Name of Recipient</th><th>Position</th><th>Date</th><th>Currency</th><th>Gifts</th><th>Meals</th><th>Travel</th><th>Entertainment</th><th>Others</th><th>Total</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <p class="small muted" style="margin-top:10px;"><i>This history is read-only — it helps you sense-check a recipient before proposing new spending.</i></p>
    </div>`;
  },

  openSixMonthSummary() {
    this.state.modal = { type: "sixMonthSummary", recipients: this._tabCRecipients || [] };
    this.render();
  },

  renderTabD(w) {
    const open = this._transTypeOpen;
    return `
    <div class="card">
      <div class="field-grid">
        <div class="field">
          <label>Type of Proposed Transaction <span class="req">*</span></label>
          <div class="combo">
            <input type="text" readonly value="${w.transactionTypes.length ? w.transactionTypes.join(", ") : ""}" placeholder="Select a transaction type" onclick="App.toggleTransTypeDropdown()">
            <div id="transTypeList" class="combo-list" style="display:${open ? "block" : "none"};">
              ${TRANSACTION_TYPES.map(t => `<label style="display:flex;align-items:center;gap:8px;padding:9px 12px;cursor:pointer;">
                <input type="checkbox" ${w.transactionTypes.includes(t) ? "checked" : ""} onchange="App.toggleTransactionType('${t}')"> ${t}
              </label>`).join("")}
            </div>
          </div>
          <div class="chips">${w.transactionTypes.map(t => `<span class="chip">${t}</span>`).join("")}</div>
        </div>
        <div class="field full">
          <label>Detailed Description of the Proposed Transaction</label>
          <textarea oninput="App.state.wizard.description=this.value">${esc(w.description)}</textarea>
        </div>
      </div>
    </div>`;
  },
  toggleTransTypeDropdown() {
    this._transTypeOpen = !this._transTypeOpen;
    this.render();
  },
  toggleTransactionType(t) {
    const w = this.state.wizard;
    const i = w.transactionTypes.indexOf(t);
    if (i >= 0) {
      w.transactionTypes.splice(i, 1);
      (TYPE_TO_AMOUNT_FIELDS[t] || []).forEach(f => { w.amounts[f] = f === "othersLabel" ? "" : 0; });
    } else {
      w.transactionTypes.push(t);
    }
    this._transTypeOpen = true;
    this.render();
  },

  unlockedFields(w) {
    const set = new Set();
    w.transactionTypes.forEach(t => (TYPE_TO_AMOUNT_FIELDS[t] || []).forEach(f => set.add(f)));
    return set;
  },
  wizardTotal(w) {
    const unlocked = this.unlockedFields(w);
    let sum = 0;
    AMOUNT_FIELDS.forEach(f => { if (unlocked.has(f.key)) sum += Number(w.amounts[f.key]) || 0; });
    if (unlocked.has("othersAmount")) sum += Number(w.amounts.othersAmount) || 0;
    return sum;
  },

  renderTabE(w) {
    const unlocked = this.unlockedFields(w);
    const amtInput = (f) => `<input type="number" min="0" step="0.01" ${unlocked.has(f.key) ? "" : "disabled"}
        value="${w.amounts[f.key]}" oninput="App.state.wizard.amounts['${f.key}']=parseFloat(this.value)||0;App.updateTotalDisplay();">`;
    return `
    <div class="card">
      <div class="field-grid" style="margin-bottom:6px;">
        <div class="field"><h3 style="margin:0;">Amount (${esc(w.currency)})</h3></div>
        <div class="field">
          <label style="text-align:right;">Currency <span class="req">*</span></label>
          <select onchange="App.state.wizard.currency=this.value;App.render();">
            <option value="">Select a currency</option>
            ${CURRENCIES.map(c => `<option value="${c}" ${w.currency === c ? "selected" : ""}>${c}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="field-grid">
        ${AMOUNT_FIELDS.map(f => `<div class="field"><label>${f.label}</label>${amtInput(f)}</div>`).join("")}
        <div class="field">
          <label>Others - please state</label>
          <input type="text" ${unlocked.has("othersLabel") ? "" : "disabled"} value="${esc(w.amounts.othersLabel)}" oninput="App.state.wizard.amounts.othersLabel=this.value">
        </div>
        <div class="field">
          <label>Others - Amount</label>
          <input type="number" min="0" step="0.01" ${unlocked.has("othersAmount") ? "" : "disabled"} value="${w.amounts.othersAmount}"
            oninput="App.state.wizard.amounts.othersAmount=parseFloat(this.value)||0;App.updateTotalDisplay();">
        </div>
      </div>
      <div class="total-row">Total Amount : <span id="totalAmountDisplay">${fmtMoney(this.wizardTotal(w))}</span></div>
      <div class="field full" style="margin-top:16px;">
        <label>State how and to whom payments will be made</label>
        <textarea oninput="App.state.wizard.paymentTo=this.value">${esc(w.paymentTo)}</textarea>
      </div>
    </div>`;
  },
  updateTotalDisplay() {
    const el = document.getElementById("totalAmountDisplay");
    if (el) el.textContent = fmtMoney(this.wizardTotal(this.state.wizard));
  },

  renderTabF(w) {
    return `
    <div class="card">
      <div class="field full">
        <label>Remarks / Comments</label>
        <textarea oninput="App.state.wizard.remarks=this.value">${esc(w.remarks)}</textarea>
      </div>
    </div>`;
  },

  openDHModal() {
    const w = this.state.wizard;
    for (const t of ["A", "B", "D", "E"]) {
      const err = this.wizardValidateTab(t);
      if (err) { this.toast(err); this.setWizardTab(t); return; }
    }
    /* Higher Management requestors skip the DH stage entirely (their chain
       starts at CFO) — per the Pre-Approval Form Approval Flow chart, so
       there's nothing to pick here; submit straight away. */
    if (isHigherManagement(w.requestor)) { this.finalizePreApproval(""); return; }
    this.state.modal = { type: "selectDH", value: w.departmentHead || "" };
    this.render();
  },
  submitPreApproval() {
    const val = this.state.modal.value;
    if (!val) { this.toast("Please select your department head."); return; }
    this.finalizePreApproval(val);
  },
  async finalizePreApproval(dh) {
    const w = this.state.wizard;
    w.departmentHead = dh;
    const payload = {
      submittedBy: this.state.session.employee.name,
      requestor: w.requestor, recipients: w.recipients, transactionTypes: w.transactionTypes,
      description: w.description, currency: w.currency, amounts: w.amounts,
      paymentTo: w.paymentTo, remarks: w.remarks, departmentHead: w.departmentHead
    };
    const rec = await Store.createPreApproval(payload);
    this.state.modal = null;
    this.state.wizard = null;
    const syncMsg = rec.syncError ? " (saved locally — SharePoint sync failed, see console)" : "";
    const notifyMsg = rec.notify ? ` Notifying ${rec.notify.name} (${rec.notify.title}) → ${rec.notify.email}` : "";
    this.toast("Submitted for approval — " + rec.refNo + "." + notifyMsg + syncMsg);
    this.openSubmission("preapproval");
  },

  /* ======================= SUBMISSION (Check My Submission) ======================= */

  openSubmission(tab) {
    this.state.sub = { tab: tab || "preapproval", ref: "", requestorName: "All", preApprovalStatus: "All", claimStatus: "All" };
    this.state.detail = null;
    this.state.view = "submission";
    this.render();
  },
  setSubTab(tab) {
    this.state.sub.tab = tab;
    this.render();
  },

  myRecords() {
    const me = this.state.session.employee.name;
    return Store.all().filter(r => r.submittedBy === me);
  },
  bucket(rec) {
    if (rec.status === "Rejected" || rec.status === "Cancelled") return "history";
    if (rec.status === "Pending") return "preapproval";
    if (rec.status === "Approved") {
      if (!rec.claim || rec.claim.status === "Open") return "claim";
      return "history";
    }
    return "history";
  },

  renderSubmission() {
    const sub = this.state.sub;
    const recs = this.myRecords();
    const requestorNames = ["All", ...Array.from(new Set(recs.map(r => r.requestor.name)))];
    return `
    <div class="page">
      <div class="page-header">
        <h2>Submission</h2>
        <button class="btn btn-secondary" onclick="App.goHome()">Back</button>
      </div>
      <div class="subtabs tabs">
        <a class="tab ${sub.tab === "preapproval" ? "active" : ""}" onclick="App.setSubTab('preapproval')">ABC Pre-Approval Form</a>
        <a class="tab ${sub.tab === "claim" ? "active" : ""}" onclick="App.setSubTab('claim')">Claim Form</a>
        <a class="tab ${sub.tab === "history" ? "active" : ""}" onclick="App.setSubTab('history')">Claim History</a>
      </div>
      ${sub.tab === "preapproval" ? this.renderPreApprovalList(recs, requestorNames) : ""}
      ${sub.tab === "claim" ? this.renderClaimList(recs, requestorNames) : ""}
      ${sub.tab === "history" ? this.renderHistoryList(recs, requestorNames) : ""}
    </div>`;
  },

  requestorFilterRow(sub, requestorNames) {
    const emp = requestorNames.length && sub.requestorName !== "All"
      ? (EMPLOYEES.find(e => e.name === sub.requestorName) || {}).employeeNo : "";
    return `
    <div class="filters">
      <div class="field"><label>Search ABC Ref No</label><input type="text" value="${esc(sub.ref)}" oninput="App.state.sub.ref=this.value;App.render();"></div>
      <div class="field">
        <label>Name of Requestor</label>
        <select onchange="App.state.sub.requestorName=this.value;App.render();">
          ${requestorNames.map(n => `<option value="${esc(n)}" ${sub.requestorName === n ? "selected" : ""}>${esc(n)}</option>`).join("")}
        </select>
      </div>
      ${sub.requestorName !== "All" ? `<div class="field"><label>Employee No</label><input type="text" value="${esc(emp || "")}" disabled></div>` : ""}
    </div>`;
  },

  filteredRecs(bucketName) {
    const sub = this.state.sub;
    let recs = this.myRecords().filter(r => this.bucket(r) === bucketName);
    if (sub.ref) recs = recs.filter(r => r.refNo.toLowerCase().includes(sub.ref.toLowerCase()));
    if (sub.requestorName && sub.requestorName !== "All") recs = recs.filter(r => r.requestor.name === sub.requestorName);
    return recs;
  },

  renderPreApprovalList(all, requestorNames) {
    const sub = this.state.sub;
    const recs = this.filteredRecs("preapproval");
    const rows = recs.map(r => `
      <tr>
        <td><a class="ref-link" onclick="App.openDetail('${r.id}','preapproval')">${esc(r.refNo)}</a></td>
        <td>${esc(r.requestor.name)}</td>
        <td>${esc(r.submittedBy || r.requestor.name)}</td>
        <td>${fmtDate(r.dateSubmitted)}</td>
        <td><span class="pill ${pillClass(r.status)}">${esc(r.status)}</span></td>
        <td>
          <button class="btn-icon" title="View" onclick="App.openDetail('${r.id}','preapproval')">&#128065;</button>
          ${(r.status === "Pending" || r.status === "Approved") ? `<button class="btn-icon" title="Cancel / Withdraw" onclick="App.openCancelModal('${r.id}')">&#10005;</button>` : ""}
        </td>
      </tr>`).join("");
    return `
    ${this.requestorFilterRow(sub, requestorNames)}
    <div class="card" style="padding:0;">
      <div class="table-wrap">
        <table class="data">
          <thead><tr><th>ABC Reference No</th><th>Name of Requestor</th><th>Submitted By</th><th>Date Submitted</th><th>Pre-Approval Status</th><th>Action</th></tr></thead>
          <tbody>${rows || `<tr class="empty-row"><td colspan="6">No pre-approval submissions found.</td></tr>`}</tbody>
        </table>
      </div>
    </div>`;
  },

  renderClaimList(all, requestorNames) {
    const sub = this.state.sub;
    let recs = this.filteredRecs("claim");
    if (sub.claimStatus && sub.claimStatus !== "All") recs = recs.filter(r => (r.claim ? r.claim.status : "Open") === sub.claimStatus);
    const rows = recs.map(r => {
      const cs = r.claim ? r.claim.status : "Open";
      return `<tr>
        <td><a class="ref-link" onclick="App.openDetail('${r.id}','claim')">${esc(r.refNo)}</a></td>
        <td>${esc(r.requestor.name)}</td>
        <td>${esc(r.submittedBy || r.requestor.name)}</td>
        <td>${fmtDate(r.dateSubmitted)}</td>
        <td><span class="pill ${pillClass(cs)}">${esc(cs)}</span></td>
        <td><button class="btn-icon" title="Open" onclick="App.openDetail('${r.id}','claim')">&#128065;</button></td>
      </tr>`;
    }).join("");
    return `
    ${this.requestorFilterRow(sub, requestorNames)}
    <div class="filters" style="margin-top:-14px;">
      <div class="field"><label>Claim Status</label>
        <select onchange="App.state.sub.claimStatus=this.value;App.render();">
          ${["All", "Open", "Closed", "Rejected", "Cancelled"].map(s => `<option ${sub.claimStatus === s ? "selected" : ""}>${s}</option>`).join("")}
        </select>
      </div>
    </div>
    <div class="card" style="padding:0;">
      <div class="table-wrap">
        <table class="data">
          <thead><tr><th>ABC Reference No</th><th>Name of Requestor</th><th>Submitted By</th><th>Date Submitted</th><th>Claim Status</th><th>Action</th></tr></thead>
          <tbody>${rows || `<tr class="empty-row"><td colspan="6">No claims to work on yet — approved pre-approvals appear here.</td></tr>`}</tbody>
        </table>
      </div>
    </div>`;
  },

  renderHistoryList(all, requestorNames) {
    const sub = this.state.sub;
    let recs = this.filteredRecs("history");
    if (sub.preApprovalStatus && sub.preApprovalStatus !== "All") recs = recs.filter(r => r.status === sub.preApprovalStatus);
    if (sub.claimStatus && sub.claimStatus !== "All") recs = recs.filter(r => (r.claim ? r.claim.status : "-") === sub.claimStatus);
    const rows = recs.map(r => `
      <tr>
        <td><a class="ref-link" onclick="App.openDetail('${r.id}','history')">${esc(r.refNo)}</a></td>
        <td>${esc(r.requestor.name)}</td>
        <td>${esc(r.submittedBy || r.requestor.name)}</td>
        <td>${fmtDate(r.dateSubmitted)}</td>
        <td><span class="pill ${pillClass(r.status)}">${esc(r.status)}</span></td>
        <td><span class="pill ${pillClass(r.claim ? r.claim.status : "-")}">${esc(r.claim ? r.claim.status : "-")}</span></td>
        <td><button class="btn-icon" title="View" onclick="App.openDetail('${r.id}','history')">&#128065;</button></td>
      </tr>`).join("");
    return `
    ${this.requestorFilterRow(sub, requestorNames)}
    <div class="filters" style="margin-top:-14px;">
      <div class="field"><label>Pre-Approval Status</label>
        <select onchange="App.state.sub.preApprovalStatus=this.value;App.render();">
          ${["All", "Approved", "Rejected", "Cancelled"].map(s => `<option ${sub.preApprovalStatus === s ? "selected" : ""}>${s}</option>`).join("")}
        </select>
      </div>
      <div class="field"><label>Claim Status</label>
        <select onchange="App.state.sub.claimStatus=this.value;App.render();">
          ${["All", "Closed", "Rejected", "Cancelled", "-"].map(s => `<option ${sub.claimStatus === s ? "selected" : ""}>${s}</option>`).join("")}
        </select>
      </div>
    </div>
    <div class="card" style="padding:0;">
      <div class="table-wrap">
        <table class="data">
          <thead><tr><th>ABC Reference No</th><th>Name of Requestor</th><th>Submitted By</th><th>Date Submitted</th><th>Pre-Approval Status</th><th>Claim Status</th><th>Action</th></tr></thead>
          <tbody>${rows || `<tr class="empty-row"><td colspan="7">No closed or rejected submissions yet.</td></tr>`}</tbody>
        </table>
      </div>
    </div>`;
  },

  openCancelModal(id) {
    this.state.modal = { type: "cancel", id, reason: "" };
    this.render();
  },
  confirmCancel() {
    const { id, reason } = this.state.modal;
    Store.cancelPreApproval(id, reason);
    this.state.modal = null;
    this.toast("Submission cancelled / withdrawn.");
    this.render();
  },

  /* ======================= DETAIL VIEW (shared: staff review + approver review) ======================= */

  openDetail(id, fromTab) {
    let rec = Store.get(id);
    if (!rec) return;
    if (fromTab === "claim" && rec.status === "Approved" && !rec.claim) {
      Store.ensureClaim(id);
      rec = Store.get(id);
    }
    this.state.detail = {
      id, rail: fromTab === "claim" && rec.claim ? "claim" : "preapproval", paTab: "A", claimSub: "form",
      approvalWhich: rec.claim ? "claim" : "preapproval",
      role: "staff", editable: !!(rec.claim && rec.claim.status === "Open" && !rec.claim.submittedDate),
      returnTo: { view: "submission" }
    };
    this.state.view = "detail";
    this.render();
  },

  closeDetail() {
    if (this.state.detail && this.state.detail.returnTo && this.state.detail.returnTo.view === "approverHome") {
      this.state.view = "approverHome";
    } else {
      this.state.view = "submission";
    }
    this.state.detail = null;
    this.render();
  },

  setRail(r) {
    this.state.detail.rail = r;
    this.render();
  },
  setDetailPATab(t) {
    this.state.detail.paTab = t;
    this.render();
  },
  setDetailClaimSub(t) {
    this.state.detail.claimSub = t;
    this.render();
  },
  setApprovalWhich(w) {
    this.state.detail.approvalWhich = w;
    this.render();
  },

  renderDetailPage() {
    const d = this.state.detail;
    const rec = Store.get(d.id);
    if (!rec) return `<div class="page"><p>Record not found.</p></div>`;
    return `
    <div class="page">
      <div class="page-header">
        <h2>${d.rail === "claim" ? "Claim Form" : d.rail === "approval" ? "Approval Status" : "Gift, Meal, Travel & Entertainment Pre-Approval Form"}
          ${d.rail === "preapproval" ? `<span class="pill ${pillClass(rec.status)}">Pre-Approval ${esc(rec.status)}</span>` : ""}
          ${d.rail === "claim" && rec.claim ? `<span class="pill ${pillClass(rec.claim.status)}">Claim ${esc(rec.claim.status)}</span>` : ""}
        </h2>
        <button class="btn btn-secondary" onclick="App.closeDetail()">Back</button>
      </div>
      <div class="detail-layout">
        <div class="rail">
          <button class="${d.rail === "preapproval" ? "active" : ""}" title="Pre-Approval Form" onclick="App.setRail('preapproval')">&#128203;</button>
          <button class="${d.rail === "claim" ? "active" : ""}" title="Claim Form" ${rec.claim ? "" : "disabled"} onclick="App.setRail('claim')">&#128196;</button>
          <button class="${d.rail === "approval" ? "active" : ""}" title="Approval Status" onclick="App.setRail('approval')">&#9989;</button>
        </div>
        <div class="detail-main">
          ${d.rail === "preapproval" ? this.renderDetailPreApproval(rec) : ""}
          ${d.rail === "claim" ? this.renderDetailClaim(rec) : ""}
          ${d.rail === "approval" ? this.renderDetailApproval(rec) : ""}
        </div>
      </div>
    </div>`;
  },

  renderDetailPreApproval(rec) {
    const d = this.state.detail;
    const tabs = [["A", "Requestor Info"], ["B", "Recipient(s) Info"], ["C", "6-Month Record"], ["D", "Transaction Details"], ["E", "Payment Amounts"], ["F", "Remarks"]];
    let body = "";
    if (d.paTab === "A") {
      const onBehalf = rec.submittedBy && rec.submittedBy !== rec.requestor.name;
      body = `<div class="card"><div class="field-grid">
        <div class="field"><label>Name of Requestor</label><div class="readonly-box">${esc(rec.requestor.name)}</div></div>
        <div class="field"><label>Employee No</label><div class="readonly-box">${esc(rec.requestor.employeeNo)}</div></div>
        <div class="field"><label>Department</label><div class="readonly-box">${esc(rec.requestor.department)}</div></div>
        <div class="field"><label>Position</label><div class="readonly-box">${esc(rec.requestor.position)}</div></div>
        <div class="field"><label>Submitted By</label><div class="readonly-box">${esc(rec.submittedBy || rec.requestor.name)}</div></div>
      </div>
      ${onBehalf ? `<p class="small muted" style="margin-top:12px;">Submitted on behalf of ${esc(rec.requestor.name)} by ${esc(rec.submittedBy)}.</p>` : ""}
      </div>`;
    } else if (d.paTab === "B") {
      body = `<div class="card"><div class="table-wrap"><table class="data">
        <thead><tr><th>No</th><th>Name of Recipient</th><th>Position</th><th>Company / Organization</th><th>Relationship</th><th>Official</th></tr></thead>
        <tbody>${rec.recipients.map((r, i) => `<tr><td>${i + 1}</td><td>${esc(r.name)}</td><td>${esc(r.position)}</td><td>${esc(r.company)}</td><td>${esc(r.relationship || "-")}</td><td>${esc(r.isOfficial || "No")}</td></tr>`).join("")}</tbody>
      </table></div></div>`;
    } else if (d.paTab === "C") {
      body = this.renderTabC(rec.recipients);
    } else if (d.paTab === "D") {
      body = `<div class="card"><div class="field-grid">
        <div class="field"><label>Type of Proposed Transaction</label><div class="chips">${rec.transactionTypes.map(t => `<span class="chip">${t}</span>`).join("")}</div></div>
        <div class="field full"><label>Detailed Description</label><div class="readonly-box">${esc(rec.description || "-")}</div></div>
      </div></div>`;
    } else if (d.paTab === "E") {
      const unlocked = new Set();
      rec.transactionTypes.forEach(t => (TYPE_TO_AMOUNT_FIELDS[t] || []).forEach(f => unlocked.add(f)));
      body = `<div class="card">
        <div class="field-grid" style="margin-bottom:6px;">
          <div class="field"><h3 style="margin:0;">Amount (${esc(rec.currency)})</h3></div>
          <div class="field"><label style="text-align:right;">Currency</label><div class="readonly-box">${esc(rec.currency)}</div></div>
        </div>
        <div class="field-grid">
          ${AMOUNT_FIELDS.map(f => `<div class="field"><label>${f.label}</label><div class="readonly-box">${unlocked.has(f.key) ? fmtMoney(rec.amounts[f.key]) : "-"}</div></div>`).join("")}
          <div class="field"><label>Others - please state</label><div class="readonly-box">${esc(rec.amounts.othersLabel || "-")}</div></div>
          <div class="field"><label>Others - Amount</label><div class="readonly-box">${unlocked.has("othersAmount") ? fmtMoney(rec.amounts.othersAmount) : "-"}</div></div>
        </div>
        <div class="total-row">Total Amount : ${fmtMoney(computePreApprovalTotalUSD(rec) === 0 ? 0 : (rec.currency ? AMOUNT_FIELDS.reduce((s, f) => s + (Number(rec.amounts[f.key]) || 0), 0) + (Number(rec.amounts.othersAmount) || 0) : 0))} <span class="muted small">(&asymp; USD ${fmtMoney(computePreApprovalTotalUSD(rec))})</span></div>
        <div class="field full" style="margin-top:16px;"><label>State how and to whom payments will be made</label><div class="readonly-box">${esc(rec.paymentTo || "-")}</div></div>
      </div>`;
    } else if (d.paTab === "F") {
      body = `<div class="card"><div class="field full"><label>Remarks / Comments</label><div class="readonly-box">${esc(rec.remarks || "-")}</div></div>
        <div class="field" style="margin-top:16px;"><label>Department Head</label><div class="readonly-box">${esc(rec.departmentHead) || "N/A — Higher Management requestor, chain starts at CFO"}</div></div>
        ${rec.status === "Cancelled" ? `<div class="banner warn" style="margin-top:16px;"><span>&#10005;</span><div><b>Cancelled / Withdrawn</b>${esc(rec.cancelReason || "No reason given.")} &middot; ${fmtDate(rec.cancelDate)}</div></div>` : ""}
      </div>`;
    }
    return `
      <div class="subtabs tabs">${tabs.map(([k, l]) => `<a class="tab ${d.paTab === k ? "active" : ""}"><span class="letter">${k}</span><span onclick="App.setDetailPATab('${k}')">${l}</span></a>`).join("")}</div>
      ${body}`;
  },

  renderDetailClaim(rec) {
    const d = this.state.detail;
    if (!rec.claim) return `<div class="card"><p class="muted">No claim yet.</p></div>`;
    const c = rec.claim;
    const editable = d.editable && c.status === "Open" && !c.submittedDate;
    const tabs = [["form", "Claim Form"], ["list", "Claim Form List View"], ["register", "Claim Register"]];
    let body = "";
    if (d.claimSub === "form") {
      body = `<div class="card"><div class="field-grid">
        <div class="field"><label>ABC No</label><div class="readonly-box">${esc(rec.refNo)}</div></div>
        <div class="field"><label>Position</label><div class="readonly-box">${esc(rec.requestor.position)}</div></div>
        <div class="field"><label>Name</label><div class="readonly-box">${esc(rec.requestor.name)}</div></div>
        <div class="field"><label>Submitted By</label><div class="readonly-box">${esc(rec.submittedBy || rec.requestor.name)}</div></div>
        <div class="field">
          <label>Cost To Be Bear By Which Company <span class="req">*</span></label>
          ${editable ? `<select onchange="App.saveClaimField('costBearBy',this.value)"><option value="">Select a company</option>${COMPANIES.map(co => `<option ${c.costBearBy === co ? "selected" : ""}>${co}</option>`).join("")}</select>` : `<div class="readonly-box">${esc(c.costBearBy || "-")}</div>`}
        </div>
        <div class="field">
          <label>Company</label>
          ${editable ? `<select onchange="App.saveClaimField('company',this.value)"><option value="">Select a company</option>${COMPANIES.map(co => `<option ${c.company === co ? "selected" : ""}>${co}</option>`).join("")}</select>` : `<div class="readonly-box">${esc(c.company || "-")}</div>`}
        </div>
        <div class="field"><label>Department</label><div class="readonly-box">${esc(rec.requestor.department)}</div></div>
        <div class="field">
          <label>Conversion Rate <span class="req">*</span> <span class="sub-label">(USD to ${esc(rec.currency || "MYR")})</span></label>
          ${editable ? `<input type="number" step="0.0001" value="${c.conversionRate}" oninput="App.state.detail._rate=this.value" onblur="App.saveClaimField('conversionRate',parseFloat(this.value)||0)">` : `<div class="readonly-box">${esc(c.conversionRate || "-")}</div>`}
        </div>
        <div class="field"><label>Date <span class="req">*</span></label>
          ${editable ? `<input type="date" value="${esc(c.date)}" onchange="App.saveClaimField('date',this.value)">` : `<div class="readonly-box">${fmtDate(c.date)}</div>`}
        </div>
        <div class="field full"><label>Purpose of Claim <span class="req">*</span></label>
          ${editable ? `<textarea oninput="App.state.detail._purpose=this.value" onblur="App.saveClaimField('purposeOfClaim',this.value)">${esc(c.purposeOfClaim)}</textarea>` : `<div class="readonly-box">${esc(c.purposeOfClaim || "-")}</div>`}
        </div>
      </div></div>`;
    } else if (d.claimSub === "list") {
      const totalMYR = claimActualMYR(c);
      const totalUSD = claimActualUSD(c);
      const approvedBudget = computePreApprovalTotalUSD(rec);
      body = `<div class="card">
        <div class="table-wrap"><table class="data">
          <thead><tr><th>No</th><th>Date</th><th>Description</th><th>Has Receipt</th><th>Transaction Type</th><th>Payment Method</th><th>Amount (MYR)</th><th>Rate (USD to MYR)</th><th>Purpose</th>${editable ? "<th></th>" : ""}</tr></thead>
          <tbody>
          ${c.lineItems.map((li, i) => `<tr>
            <td>${i + 1}</td><td>${fmtDate(li.date)}</td><td>${esc(li.description)}</td>
            <td>${esc(li.hasReceipt)}</td><td>${esc(li.transactionType)}</td><td>${esc(li.paymentMethod)}</td>
            <td>${fmtMoney(li.amountMYR)}</td><td>${li.rate}</td><td>${esc(li.purpose)}</td>
            ${editable ? `<td><button class="btn-icon" onclick="App.removeClaimLineItem('${li.id}')">&#128465;</button></td>` : ""}
          </tr>`).join("") || `<tr class="empty-row"><td colspan="${editable ? 10 : 9}">No line items yet.</td></tr>`}
          </tbody>
        </table></div>
        ${editable ? `<div class="icon-btn-row"><button class="fab" title="Add line item" onclick="App.openAddLineItemModal()">+</button></div>` : ""}
        <div class="field full" style="margin-top:10px;">
          <label>Attachments</label>
          <div class="readonly-box">${c.attachments && c.attachments.length ? c.attachments.map(a => esc(a)).join(", ") : "There is nothing attached."}</div>
          ${editable ? `<button class="btn btn-secondary btn-sm" style="margin-top:8px;width:fit-content;" onclick="App.attachFile()">&#128206; Attach file</button>` : ""}
        </div>
        <div class="total-row" style="flex-direction:column;align-items:flex-end;gap:4px;">
          <div>Actual Spending Amount in MYR : ${fmtMoney(totalMYR)}</div>
          <div>Actual Spending Amount in USD : ${fmtMoney(totalUSD)}</div>
          <div>Approved Budgeted Amount (USD) : ${fmtMoney(approvedBudget)}</div>
        </div>
      </div>`;
    } else if (d.claimSub === "register") {
      const totalUSD = claimActualUSD(c);
      body = `<div class="card">
        <div class="field-grid" style="margin-bottom:16px;max-width:520px;">
          <div class="field"><label>Month <span class="req">*</span></label>
            ${editable ? `<select onchange="App.saveRegisterField('month',this.value)"><option value="">Select</option>${["January","February","March","April","May","June","July","August","September","October","November","December"].map(m => `<option ${c.register.month === m ? "selected" : ""}>${m}</option>`).join("")}</select>` : `<div class="readonly-box">${esc(c.register.month || "-")}</div>`}
          </div>
          <div class="field"><label>Year <span class="req">*</span></label>
            ${editable ? `<input type="number" value="${c.register.year}" onchange="App.saveRegisterField('year',parseInt(this.value)||c.register.year)">` : `<div class="readonly-box">${esc(String(c.register.year || "-"))}</div>`}
          </div>
        </div>
        <div class="section-title">Total (USD) : ${fmtMoney(totalUSD)}
          ${editable ? `<a onclick="App.openAddRecipientModal(true)">Can't find the recipient in the dropdown list? Click here</a>` : ""}
        </div>
        <div class="table-wrap"><table class="data">
          <thead><tr><th>No</th><th>Date</th><th>Name of Recipient</th><th>Recipient's Company</th><th>Others</th><th>Is Official</th>${editable ? "<th></th>" : ""}</tr></thead>
          <tbody>${c.register.recipients.map((r, i) => `<tr>
            <td>${i + 1}</td><td>${fmtDate(r.date)}</td><td>${esc(r.name)}</td><td>${esc(r.company)}</td><td>${fmtMoney(r.others)}</td><td>${esc(r.isOfficial)}</td>
            ${editable ? `<td><button class="btn-icon" onclick="App.removeRegisterRecipient(${i})">&#128465;</button></td>` : ""}
          </tr>`).join("") || `<tr class="empty-row"><td colspan="${editable ? 7 : 6}">No recipients.</td></tr>`}</tbody>
        </table></div>
        ${editable && !c.submittedDate ? `<div class="icon-btn-row"><button class="btn btn-primary" onclick="App.submitClaimRegister()">Submit</button></div>` : ""}
        ${c.submittedDate ? `<p class="small muted" style="text-align:right;margin-top:10px;">Submitted ${fmtDate(c.submittedDate)} &mdash; view approval status via the left-panel tick icon.</p>` : ""}
      </div>`;
    }
    return `
      <div class="subtabs tabs">${tabs.map(([k, l]) => `<a class="tab ${d.claimSub === k ? "active" : ""}" onclick="App.setDetailClaimSub('${k}')">${l}</a>`).join("")}</div>
      ${body}`;
  },

  saveClaimField(field, value) {
    Store.saveClaimForm(this.state.detail.id, { [field]: value });
    this.render();
  },
  saveRegisterField(field, value) {
    Store.saveClaimRegister(this.state.detail.id, { [field]: value });
    this.render();
  },
  removeClaimLineItem(itemId) {
    Store.removeClaimLineItem(this.state.detail.id, itemId);
    this.render();
  },
  attachFile() {
    const rec = Store.get(this.state.detail.id);
    const name = "receipt-" + (rec.claim.attachments.length + 1) + ".pdf";
    if (!rec.claim.attachments) rec.claim.attachments = [];
    rec.claim.attachments.push(name);
    Store.save();
    this.toast("Attached " + name + " (demo — no real upload).");
    this.render();
  },
  removeRegisterRecipient(i) {
    Store.removeRegisterRecipient(this.state.detail.id, i);
    this.render();
  },
  submitClaimRegister() {
    const rec = Store.get(this.state.detail.id);
    if (!rec.claim.register.month) { this.toast("Please select the claim's Month and Year first."); return; }
    if (!rec.claim.costBearBy || !rec.claim.conversionRate || !rec.claim.date || !rec.claim.purposeOfClaim) { this.toast("Please complete the Claim Form tab first."); return; }
    if (!rec.claim.lineItems.length) { this.toast("Please add at least one claim line item."); return; }
    const notifyTarget = Store.submitClaim(this.state.detail.id);
    this.state.detail.editable = false;
    this.state.detail.rail = "approval";
    this.state.detail.approvalWhich = "claim";
    const notifyMsg = notifyTarget ? ` Notifying ${notifyTarget.name} (${notifyTarget.title}) → ${notifyTarget.email}` : "";
    this.toast("Claim submitted for approval." + notifyMsg);
    this.render();
  },

  openAddLineItemModal() {
    this.state.modal = { type: "addLineItem", form: { date: todayISO(), description: "", hasReceipt: "Yes", transactionType: TRANSACTION_TYPES[0], paymentMethod: PAYMENT_METHODS[0], amountMYR: 0, rate: 4.2, purpose: "" } };
    this.render();
  },
  submitAddLineItemModal() {
    const f = this.state.modal.form;
    Store.addClaimLineItem(this.state.detail.id, f);
    this.state.modal = null;
    this.render();
  },

  renderDetailApproval(rec) {
    const d = this.state.detail;
    const which = rec.claim ? d.approvalWhich : "preapproval";
    const stages = which === "claim" ? (rec.claim ? rec.claim.approvals : []) : rec.approvals;
    const acting = this._actingStage(rec, which);
    return `
      ${rec.claim ? `<div class="tabs" style="margin-bottom:14px;">
        <a class="tab ${which === "preapproval" ? "active" : ""}" onclick="App.setApprovalWhich('preapproval')">Pre-Approval Form</a>
        <a class="tab ${which === "claim" ? "active" : ""}" onclick="App.setApprovalWhich('claim')">Claim Form</a>
      </div>` : ""}
      <div class="card">
        <div class="table-wrap"><table class="data">
          <thead><tr><th>No</th><th>Title</th><th>Name</th><th>Role</th><th>Date</th><th>Status</th><th>Comments</th><th>Reject Reason</th></tr></thead>
          <tbody>${stages.map((s, i) => `<tr>
            <td>${i + 1}</td><td>${esc(s.title)}</td><td>${esc(s.name || "-")}</td><td>${esc(s.role)}</td>
            <td>${s.date ? fmtDate(s.date) : "-"}</td>
            <td><span class="pill ${stagePillClass(s.status)}">${stageLabel(s.status)}</span></td>
            <td>${esc(s.comments || "-")}</td><td>${esc(s.rejectReason || "-")}</td>
          </tr>`).join("")}</tbody>
        </table></div>
      </div>
      ${acting !== null ? this.renderActionPanel(rec, which, acting) : ""}`;
  },

  /* Real-name match against the signed-in Microsoft account — this is the
     actual authorization check for approving/rejecting a stage. Anyone
     signed in can see this page (approver review is not secret), but only
     the real person named on the currently-pending stage gets the
     Approve/Reject panel. */
  _actingStage(rec, which) {
    const s = this.state.session;
    if (!s || !s.employee) return null;
    const stages = which === "claim" ? (rec.claim ? rec.claim.approvals : []) : rec.approvals;
    const idx = stages.findIndex(st => st.status === "pending" && st.name === s.employee.name);
    return idx >= 0 ? idx : null;
  },

  renderActionPanel(rec, which, stageIndex) {
    return `
    <div class="card">
      <h3>Your decision &mdash; ${esc((which === "claim" ? rec.claim.approvals : rec.approvals)[stageIndex].title)}</h3>
      <div class="field full">
        <label>Comments</label>
        <textarea id="approveComments" placeholder="Optional comments"></textarea>
      </div>
      <div class="field full" style="margin-top:10px;">
        <label>Reject Reason <span class="sub-label">(required if rejecting)</span></label>
        <input type="text" id="rejectReason" placeholder="Reason for rejection">
      </div>
      <div class="icon-btn-row" style="justify-content:flex-start;margin-top:16px;">
        <button class="btn btn-primary" onclick="App.actOnStage('${which}','${rec.id}',${stageIndex},'approve')">Approve</button>
        <button class="btn btn-danger" onclick="App.actOnStage('${which}','${rec.id}',${stageIndex},'reject')">Reject</button>
      </div>
    </div>`;
  },

  actOnStage(which, id, stageIndex, action) {
    const comments = (document.getElementById("approveComments") || {}).value || "";
    const rejectReason = (document.getElementById("rejectReason") || {}).value || "";
    if (action === "reject" && !rejectReason) { this.toast("Please provide a reject reason."); return; }
    const notifyTarget = which === "claim"
      ? Store.actOnClaim(id, stageIndex, action, comments, rejectReason)
      : Store.actOnPreApproval(id, stageIndex, action, comments, rejectReason);
    const notifyMsg = notifyTarget ? ` Notifying ${notifyTarget.name} (${notifyTarget.title}) → ${notifyTarget.email}` : "";
    this.toast((action === "approve" ? "Approved." : "Rejected.") + notifyMsg);
    this.state.view = "approverHome";
    this.state.detail = null;
    this.render();
  },

  /* ======================= APPROVER CONSOLE ======================= */

  renderApproverHome() {
    if (!this.isApproverUser()) {
      // Defensive: bounce anyone who isn't in the approver roster back home,
      // in case state.view was ever set to "approverHome" some other way.
      this.state.view = "home";
      this.render();
      return "";
    }
    const identity = this.approverIdentity();
    const queue = Store.approverQueue(identity);
    const rows = queue.map(item => {
      const totalUSD = item.flow === "claim" ? claimActualUSD(item.rec.claim) : computePreApprovalTotalUSD(item.rec);
      return `<tr>
        <td>${esc(item.rec.refNo)}</td>
        <td>${esc(item.rec.requestor.name)}</td>
        <td>${esc(item.rec.submittedBy || item.rec.requestor.name)}</td>
        <td>${item.flow === "claim" ? "Claim Form" : "Pre-Approval Form"}</td>
        <td>${esc(item.stage.title)}</td>
        <td>&asymp; USD ${fmtMoney(totalUSD)}</td>
        <td>${fmtDate(item.rec.dateSubmitted)}</td>
        <td><button class="btn btn-secondary btn-sm" onclick="App.reviewFromQueue('${item.rec.id}','${item.flow}')">Review</button></td>
      </tr>`;
    }).join("");
    return `
    <div class="page">
      <div class="page-header"><h2>Approver Console</h2></div>
      <div class="banner">
        <span>&#9989;</span>
        <div><b>Signed in as ${esc(identity.name)}</b>${esc(identity.role)} &mdash; showing every pre-approval &amp; claim currently awaiting your decision.</div>
      </div>
      <div class="card" style="padding:0;">
        <div class="table-wrap">
          <table class="data">
            <thead><tr><th>ABC Reference No</th><th>Name of Requestor</th><th>Submitted By</th><th>Flow</th><th>Stage</th><th>Amount</th><th>Date Submitted</th><th>Action</th></tr></thead>
            <tbody>${rows || `<tr class="empty-row"><td colspan="8">Nothing awaiting your approval right now.</td></tr>`}</tbody>
          </table>
        </div>
      </div>
    </div>`;
  },

  reviewFromQueue(id, flow) {
    const rec = Store.get(id);
    this.state.detail = {
      id, rail: flow === "claim" ? "claim" : "preapproval", paTab: "A", claimSub: "form",
      approvalWhich: flow, role: "approver", editable: false,
      returnTo: { view: "approverHome" }
    };
    this.state.view = "detail";
    this.render();
  },

  /* ======================= MODALS ======================= */

  closeModal() {
    this.state.modal = null;
    this.render();
  },

  renderModal() {
    const m = this.state.modal;
    if (!m) return "";
    if (m.type === "addRecipient") return this.modalAddRecipient(m);
    if (m.type === "sixMonthSummary") return this.modalSixMonthSummary(m);
    if (m.type === "selectDH") return this.modalSelectDH(m);
    if (m.type === "cancel") return this.modalCancel(m);
    if (m.type === "addLineItem") return this.modalAddLineItem(m);
    return "";
  },

  modalAddRecipient(m) {
    const f = m.form;
    return `<div class="modal-overlay" onmousedown="if(event.target===this) App.closeModal()">
      <div class="modal">
        <h3>Add new recipient</h3>
        <div class="field"><label>Name of Recipient <span class="req">*</span></label><input type="text" value="${esc(f.name)}" oninput="App.state.modal.form.name=this.value"></div>
        <div class="field"><label>Position</label><input type="text" value="${esc(f.position)}" oninput="App.state.modal.form.position=this.value"></div>
        <div class="field"><label>Company/Organization <span class="req">*</span></label><input type="text" value="${esc(f.company)}" oninput="App.state.modal.form.company=this.value"></div>
        <div class="field"><label>Relationship with RGB</label>
          <select onchange="App.state.modal.form.relationship=this.value">
            <option value="">Select</option>${RELATIONSHIP_OPTIONS.map(o => `<option ${f.relationship === o ? "selected" : ""}>${o}</option>`).join("")}
          </select>
        </div>
        <div class="field"><label>Is the recipient an official?</label>
          <select onchange="App.state.modal.form.isOfficial=this.value">
            <option value="No" ${f.isOfficial === "No" ? "selected" : ""}>No</option>
            <option value="Yes" ${f.isOfficial === "Yes" ? "selected" : ""}>Yes</option>
          </select>
        </div>
        <div class="modal-actions">
          <button class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button class="btn btn-primary" onclick="App.submitAddRecipientModal()">Add</button>
        </div>
      </div>
    </div>`;
  },

  modalSixMonthSummary(m) {
    const rows = m.recipients.map(r => {
      const hist = sixMonthHistory(r);
      const totalUSD = hist.reduce((s, row) => s + toUSD(rowTotal(row), row.currency), 0);
      const warn = totalUSD > 1000;
      return `<tr><td>${esc(r.name)}<br><span class="muted small">${esc(r.company)}</span></td><td>${esc(r.position)}</td><td>${fmtMoney(totalUSD)}</td><td>${warn ? "&#9888;&#65039;" : ""}</td></tr>`;
    }).join("");
    return `<div class="modal-overlay" onmousedown="if(event.target===this) App.closeModal()">
      <div class="modal wide">
        <h3>Sum for Previous Six Months Record of Recipient</h3>
        <div class="table-wrap"><table class="data">
          <thead><tr><th>Name of Recipient</th><th>Position</th><th>Total Expenses in USD</th><th></th></tr></thead>
          <tbody>${rows || `<tr class="empty-row"><td colspan="4">No recipients yet.</td></tr>`}</tbody>
        </table></div>
        <div class="modal-actions"><button class="btn btn-secondary" onclick="App.closeModal()">Close</button></div>
      </div>
    </div>`;
  },

  modalSelectDH(m) {
    /* Live Entra ID group members (FCPA DH), populated on sign-in — falls
       back to the static DEPARTMENT_HEADS list only if that fetch somehow
       came back empty (there's no unauthenticated demo/preview sign-in
       anymore that would skip signInWithMicrosoft() entirely). */
    const dhList = this.state.dhOptions.length ? this.state.dhOptions : DEPARTMENT_HEADS;
    return `<div class="modal-overlay" onmousedown="if(event.target===this) App.closeModal()">
      <div class="modal">
        <h3>Please select your department head:</h3>
        <div class="field">
          <select id="dhSelect" onchange="App.state.modal.value=this.value">
            <option value="">Select DH</option>
            ${dhList.map(dh => `<option value="${esc(dh.name)}" ${m.value === dh.name ? "selected" : ""}>${esc(dh.name)}${dh.team ? " (" + dh.team + ")" : ""}</option>`).join("")}
          </select>
        </div>
        <div class="modal-actions">
          <button class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button class="btn btn-primary" onclick="App.submitPreApproval()">Submit</button>
        </div>
      </div>
    </div>`;
  },

  modalCancel(m) {
    const rec = Store.get(m.id);
    return `<div class="modal-overlay" onmousedown="if(event.target===this) App.closeModal()">
      <div class="modal">
        <h3>Cancel / withdraw ${esc(rec.refNo)}?</h3>
        <p class="small muted">This stops the request from moving further through approval. This cannot be undone.</p>
        <div class="field"><label>Reason (optional)</label><textarea oninput="App.state.modal.reason=this.value"></textarea></div>
        <div class="modal-actions">
          <button class="btn btn-secondary" onclick="App.closeModal()">Never mind</button>
          <button class="btn btn-danger" onclick="App.confirmCancel()">Cancel Submission</button>
        </div>
      </div>
    </div>`;
  },

  modalAddLineItem(m) {
    const f = m.form;
    return `<div class="modal-overlay" onmousedown="if(event.target===this) App.closeModal()">
      <div class="modal">
        <h3>Add claim line item</h3>
        <div class="field"><label>Date</label><input type="date" value="${esc(f.date)}" oninput="App.state.modal.form.date=this.value"></div>
        <div class="field"><label>Description</label><input type="text" value="${esc(f.description)}" oninput="App.state.modal.form.description=this.value"></div>
        <div class="field"><label>Has Receipt</label><select onchange="App.state.modal.form.hasReceipt=this.value"><option ${f.hasReceipt === "Yes" ? "selected" : ""}>Yes</option><option ${f.hasReceipt === "No" ? "selected" : ""}>No</option></select></div>
        <div class="field"><label>Transaction Type</label><select onchange="App.state.modal.form.transactionType=this.value">${TRANSACTION_TYPES.map(t => `<option ${f.transactionType === t ? "selected" : ""}>${t}</option>`).join("")}</select></div>
        <div class="field"><label>Payment Method</label><select onchange="App.state.modal.form.paymentMethod=this.value">${PAYMENT_METHODS.map(t => `<option ${f.paymentMethod === t ? "selected" : ""}>${t}</option>`).join("")}</select></div>
        <div class="field"><label>Amount (MYR)</label><input type="number" step="0.01" value="${f.amountMYR}" oninput="App.state.modal.form.amountMYR=parseFloat(this.value)||0"></div>
        <div class="field"><label>Rate (USD to MYR)</label><input type="number" step="0.0001" value="${f.rate}" oninput="App.state.modal.form.rate=parseFloat(this.value)||0"></div>
        <div class="field"><label>Purpose</label><input type="text" value="${esc(f.purpose)}" oninput="App.state.modal.form.purpose=this.value"></div>
        <div class="modal-actions">
          <button class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button class="btn btn-primary" onclick="App.submitAddLineItemModal()">Add</button>
        </div>
      </div>
    </div>`;
  }
};

window.addEventListener("DOMContentLoaded", () => App.init());
