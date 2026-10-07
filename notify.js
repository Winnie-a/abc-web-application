/* ==========================================================================
   ABC Application — notifications (email + Microsoft Teams)
   Sent by the web app itself through Microsoft Graph, using the signed-in
   person's own account (no Power Automate):
     - email:  POST /me/sendMail
     - Teams:  a one-to-one chat with the recipient (POST /chats, then
               POST /chats/{id}/messages) — appears as a normal chat message
               from that person, with a link, not as a bot with buttons.
   Needs the delegated permissions Mail.Send, Chat.Create and ChatMessage.Send.
   They are requested SEPARATELY from sign-in (the "Enable notifications"
   button on the home screen), so a refused permission can never block sign-in.

   TESTING: set localStorage "abc_notify_redirect" to an email address and every
   message goes to that one person instead (subject tagged "[TEST ...]"), so
   no real approver gets a test message. Remove the key to go live.
   ========================================================================== */

const NOTIFY_SCOPES = ["Mail.Send", "Chat.Create", "ChatMessage.Send"];

const Notify = {
  enabled: false,
  _warned: false,

  async _token() {
    await msalReady;
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Not signed in");
    const r = await msalInstance.acquireTokenSilent({ scopes: NOTIFY_SCOPES, account });
    return r.accessToken;
  },

  // Quietly checks whether the notification permissions were already granted.
  async checkEnabled() {
    try { await this._token(); this.enabled = true; } catch (e) { this.enabled = false; }
    return this.enabled;
  },

  // Must be called from a click (opens Microsoft's consent pop-up).
  async enable() {
    await msalReady;
    await msalInstance.acquireTokenPopup({ scopes: NOTIFY_SCOPES, account: msalInstance.getActiveAccount() });
    this.enabled = true;
  },

  async _graph(path, options = {}) {
    const token = await this._token();
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...options,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }
    });
    if (!res.ok) throw new Error(`Graph ${options.method || "GET"} ${path} failed: ${res.status} ${(await res.text()).slice(0, 200)}`);
    return res.status === 202 || res.status === 204 ? null : res.json();
  },

  async sendMail(to, subject, html) {
    await this._graph("/me/sendMail", {
      method: "POST",
      body: JSON.stringify({
        message: { subject, body: { contentType: "HTML", content: html }, toRecipients: [{ emailAddress: { address: to } }] },
        saveToSentItems: false
      })
    });
  },

  // toUserId = the recipient's Entra object id; meId = the sender's.
  async sendTeams(toUserId, meId, html) {
    let chatId = "48:notes"; // the sender's own "notes to self" chat
    if (toUserId !== meId) {
      const member = id => ({ "@odata.type": "#microsoft.graph.aadUserConversationMember", roles: ["owner"], "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${id}')` });
      const chat = await this._graph("/chats", { method: "POST", body: JSON.stringify({ chatType: "oneOnOne", members: [member(meId), member(toUserId)] }) });
      chatId = chat.id;
    }
    await this._graph(`/chats/${encodeURIComponent(chatId)}/messages`, { method: "POST", body: JSON.stringify({ body: { contentType: "html", content: html } }) });
  },

  esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  },

  link(rec) {
    return `${window.location.origin}${window.location.pathname}?ref=${encodeURIComponent(rec.refNo)}`;
  },

  _userByEmail(email) {
    const e = String(email || "").trim().toLowerCase();
    const dir = (typeof App !== "undefined" && App.state && App.state.directory) || [];
    return dir.find(u => (u.email || "").toLowerCase() === e) || null;
  },

  /* Sends one notification by email AND Teams. `lines` = short facts shown in
     the message body. Never throws: resolves { mail, teams } booleans. */
  async send({ toEmail, subject, intro, lines, rec }) {
    const out = { mail: false, teams: false };
    try {
      if (!this.enabled || typeof isGraphConnected !== "function" || !isGraphConnected()) return out;
      let to = String(toEmail || "").trim();
      let tag = "";
      const redirect = (typeof localStorage !== "undefined" && localStorage.getItem("abc_notify_redirect")) || "";
      if (redirect) { tag = `[TEST — intended for ${to || "nobody"}] `; to = redirect.trim(); }
      if (!to) return out;
      const html = `<p>${intro}</p><p>${(lines || []).map(l => l).join("<br>")}</p><p><a href="${this.link(rec)}">Open ${rec.refNo} in the ABC app</a></p>`
        + (tag ? `<p><i>${tag}</i></p>` : "");
      const me = App.state.session && App.state.session.employee;
      const target = this._userByEmail(to);
      const tasks = [
        this.sendMail(to, tag + subject, html).then(() => { out.mail = true; }).catch(e => console.error("Notification email failed", e)),
        (target && me && me.graphId)
          ? this.sendTeams(target.graphId, me.graphId, `<b>${tag}${subject}</b><br>${html}`).then(() => { out.teams = true; }).catch(e => console.error("Notification Teams message failed", e))
          : Promise.resolve()
      ];
      await Promise.all(tasks);
      if (!out.mail && !out.teams && typeof App !== "undefined" && App.toast) App.toast("Saved, but the email / Teams notification could not be sent — see console.");
    } catch (e) {
      console.error("Notification failed", e);
    }
    return out;
  }
};
