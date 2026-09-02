/* ==========================================================================
   ABC Application — SharePoint sync layer
   Loads the flat SharePoint lists into the SAME nested shape Store.state
   already uses (see newBlankPreApproval() in store.js), so every existing
   *read* in app.js — Store.get(), Store.all(), the render functions — keeps
   working completely unchanged. Only Store's *write* methods need to also
   push to Graph; see actOnPreApproval's rewritten version at the bottom for
   the pattern (repeat it for the other write methods).
   ========================================================================== */

const GraphSync = {
  async loadAll() {
    const [requests, recipients, approvals, claims, lineItems, registerRecipients] = await Promise.all([
      graphListItems("ABC_Requests"),
      graphListItems("ABC_Recipients"),
      graphListItems("ABC_Approvals"),
      graphListItems("ABC_Claims"),
      graphListItems("ABC_ClaimLineItems"),
      graphListItems("ABC_ClaimRegisterRecipients")
    ]);

    const preApprovals = requests.map(req => {
      const rec = {
        id: req.id, // SharePoint item id — swap in for the uid()-generated id everywhere
        refNo: req.Title,
        submittedBy: req.SubmittedBy,
        requestor: { name: req.RequestorName, employeeNo: req.RequestorEmployeeNo, department: req.Department, position: req.Position },
        recipients: recipients.filter(r => r.RequestId === req.id)
          .map(r => ({ name: r.Name, position: r.Position, company: r.Company, relationship: r.Relationship, isOfficial: r.IsOfficial })),
        transactionTypes: (req.TransactionTypes || "").split(";").filter(Boolean),
        description: req.Description, currency: req.Currency,
        amounts: {
          gifts: req.Gifts, meals: req.Meals, entertainment: req.Entertainment,
          airfare: req.Airfare, transportation: req.Transportation, hotel: req.Hotel,
          othersLabel: req.OthersLabel, othersAmount: req.OthersAmount
        },
        paymentTo: req.PaymentTo, remarks: req.Remarks, departmentHead: req.DepartmentHead,
        dateSubmitted: req.DateSubmitted, status: req.Status,
        approvals: approvals.filter(a => a.RequestId === req.id && a.FlowType === "PreApproval")
          .sort((a, b) => a.StageOrder - b.StageOrder)
          .map(a => ({ title: a.StageTitle, role: a.Role, name: a.ApproverName, applicable: true, status: a.Status, date: a.DateActioned, comments: a.Comments, rejectReason: a.RejectReason })),
        claim: null
      };

      const claimRow = claims.find(c => c.RequestId === req.id);
      if (claimRow) {
        rec.claim = {
          company: claimRow.Company, costBearBy: claimRow.CostBearBy, date: claimRow.ClaimDate,
          conversionRate: claimRow.ConversionRate, purposeOfClaim: claimRow.PurposeOfClaim,
          lineItems: lineItems.filter(li => li.ClaimId === claimRow.id)
            .map(li => ({ id: li.id, date: li.Date, description: li.Description, hasReceipt: li.HasReceipt, transactionType: li.TransactionType, paymentMethod: li.PaymentMethod, amountMYR: li.AmountMYR, rate: li.Rate, purpose: li.Purpose, attachment: li.AttachmentLink })),
          attachments: [],
          register: {
            month: claimRow.Month, year: claimRow.Year,
            recipients: registerRecipients.filter(r => r.ClaimId === claimRow.id)
              .map(r => ({ date: r.Date, name: r.Name, company: r.Company, others: r.Others, isOfficial: r.IsOfficial }))
          },
          approvals: approvals.filter(a => a.RequestId === req.id && a.FlowType === "Claim")
            .sort((a, b) => a.StageOrder - b.StageOrder)
            .map(a => ({ title: a.StageTitle, role: a.Role, name: a.ApproverName, applicable: true, status: a.Status, date: a.DateActioned, comments: a.Comments, rejectReason: a.RejectReason })),
          status: claimRow.Status, submittedDate: claimRow.SubmittedDate
        };
      }
      return rec;
    });

    Store.state = { preApprovals, approverEmails: Store.state.approverEmails || {} };
    return Store.state;
  }
};

/* -------------------------------------------------------------------------
   Worked example: converting one write method to push to Graph.
   Apply the same shape to createPreApproval, cancelPreApproval, actOnClaim,
   addClaimLineItem, etc. — write to the list, then either patch the local
   Cache directly (optimistic, instant UI) or just re-run GraphSync.loadAll().
   ------------------------------------------------------------------------- */
Store.actOnPreApproval = async function (id, stageIndex, action, comments, rejectReason) {
  const rec = this.get(id);
  if (!rec) return null;

  advanceChain(rec.approvals, stageIndex, action, comments, rejectReason); // same local logic as before
  rec.status = chainOverallStatus(rec.approvals, "Pending", "Approved");

  // Push just the two things that changed up to SharePoint.
  const approvalListItemId = rec.approvals[stageIndex]._spItemId; // see note below
  await graphUpdateItem("ABC_Approvals", approvalListItemId, {
    Status: rec.approvals[stageIndex].status,
    DateActioned: rec.approvals[stageIndex].date,
    Comments: comments || "",
    RejectReason: rejectReason || ""
  });
  await graphUpdateItem("ABC_Requests", rec.id, { Status: rec.status });

  return action === "approve" ? pendingNotifyTarget(rec.approvals) : null;
  // Note: GraphSync.loadAll() doesn't currently keep each approval row's own
  // SharePoint item id (_spItemId) — add that field alongside the others in
  // loadAll() above once you're wiring this in for real, so this update call
  // can target the exact row.
};
