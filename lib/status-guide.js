// lib/status-guide.js — Axxerion WR/WO status reference with responsible parties
// Kept in sync with STATUS_GUIDE in public/index.html

const STATUS_GUIDE = {
  wr: [
    { status: "Draft", meaning: "Started creating but not completed", action: "Submit", party: "Creator - Any employee with access can create a request" },
    { status: "AM / DM Review", meaning: "DM has to review the work request to ensure the correct priority, scope of work is clear, work is needed", action: "GENERATE WORKORDER (Pest Control, Landscaping, Snow Removal, Junk Removal, Rekey Locks, Golf Cart Minor Repairs) or APPROVE if related to REPAIR", party: "District Manager" },
    { status: "Security Related", meaning: "Security FM will review request", action: "APPROVE if related to CCTV or Security Guard Service — Work Order Generated", party: "Security FM" },
    { status: "FM Review", meaning: "FM to review and classify correct Priority", action: "GENERATE WORKORDER to complete repair", party: "Facility Manager" },
    { status: "OLD WORK FLOW - IN PROCESS", meaning: "Pending payment of final invoice (old workflow)", action: "Work order is still open for work or invoice processing", party: "Overseas Team" },
    { status: "CLOSED (no closed date)", meaning: "Old workflow — Work order has been generated", action: "Pending WO completion / invoice paid - Link to the work order from the work request", party: "Overseas Team" },
  ],
  wo: [
    { status: "Draft", meaning: "Started creating but not completed", action: "Submit", party: "Assigner (DM, FM or Security FM)", phase: "Creation" },
    { status: "Assigned", meaning: "WO has been assigned to Vendor", action: "Accept or reject work", party: "Overseas Team", phase: "Assignment" },
    { status: "Approver Review", meaning: "Over Assigner's threshold limit - Needs approval before being assigned to vendor", action: "Next approver - APPROVE", party: "FM, Randy, Chris Runckle (based on NTE value)", phase: "Assignment" },
    { status: "Accepted", meaning: "Vendor has accepted the work", action: "Set Scheduled Start - START WORK", party: "Overseas Team", phase: "Assignment" },
    { status: "Need Reassignment", meaning: "Vendor has rejected the work", action: "Sent back to FM to assign to new vendor", party: "Overseas Team reject → FM reassigns", phase: "Assignment" },
    { status: "Work In Progress", meaning: "Work is started - All invoice fields are open to enter invoice information", action: "CHANGE ORDER submission OR END WORK to finish work", party: "Overseas Team", phase: "Execution" },
    { status: "Change Order Submitted", meaning: "Vendor costs are going to exceed NTE - Change order required before they complete work or attach ESTIMATE for change of Scope of Work", action: "SUBMIT CHANGE ORDER - FM to review change order", party: "Overseas Team", phase: "Change Orders" },
    { status: "Change Order Approver Review", meaning: "FM Reviews Change order", action: "APPROVE & REJECT - if approved, moves to next approver or vendor", party: "Assigner (DM, FM or Security FM)", phase: "Change Orders" },
    { status: "Change Order Approved", meaning: "If cost exceed Assigner's Approval Limit - routes to next approver", action: "APPROVE - moves to next approver or vendor if within threshold", party: "FM, Randy, Chris Runckle (based on NTE value)", phase: "Change Orders" },
    { status: "Change Order Rejected", meaning: "Change order has been rejected", action: "REJECT — Vendor will need to submit invoice for original NTE amount - END WORK", party: "DM, FM, Randy, Chris Runckle", phase: "Change Orders" },
    { status: "Work Finished", meaning: "Work has been completed", action: "Goes to AM to confirm work completion - APPROVED or VENDOR NOT FINISHED or REJECT WORK", party: "Area Manager", phase: "Completion" },
    { status: "Unsatisfactory", meaning: "Work has not been completed to satisfaction or as per Scope of Work", action: "Sent back to Vendor to FINISH WORK", party: "Overseas Team", phase: "Completion" },
    { status: "Prepare Financials", meaning: "Work has been completed and confirmed, but vendor has not submitted their invoice", action: "Vendor to submit invoice information - SUBMIT FINANCIALS", party: "Overseas Team", phase: "Financials" },
    { status: "Financials Submitted", meaning: "Invoice costs to be reviewed by Assigner", action: "APPROVE - marks WO as COMPLETED - Invoice created and submitted for audit review", party: "Assigner (DM, FM or Security FM)", phase: "Financials" },
    { status: "Financials Rejected", meaning: "Invoice costs have been rejected and sent back to vendor", action: "Vendor to correct and resubmit financials", party: "Assigner (DM, FM or Security FM)", phase: "Financials" },
    { status: "Invoiced", meaning: "All work has been completed, invoice submitted for payment", action: "Audit Process - Review Invoice submitted - Approve and send to Accounting for processing", party: "Axxerion Support", phase: "Audit & Close" },
    { status: "Completed", meaning: "All work has been completed, invoice submitted for payment", action: "Audit Process - Review Invoice submitted - Approve and send to Accounting for processing", party: "Axxerion Support", phase: "Audit & Close" },
  ],
};

// Flat lookup including aliases for live-data status strings
const STATUS_INFO = {};
STATUS_GUIDE.wr.forEach(s => { STATUS_INFO[s.status.toLowerCase()] = { ...s, type: "Request" }; });
STATUS_GUIDE.wo.forEach(s => { STATUS_INFO[s.status.toLowerCase()] = { ...s, type: "Work Order" }; });
// Aliases
STATUS_INFO["needs reassignment"] = STATUS_INFO["need reassignment"];
STATUS_INFO["work in progress"] = STATUS_INFO["work in progress"];
STATUS_INFO["work finished"] = STATUS_INFO["work finished"];

function getStatusInfo(statusName) {
  if (!statusName) return null;
  return STATUS_INFO[statusName.toLowerCase()] || null;
}

function listStatusesByParty(partyQuery) {
  const q = (partyQuery || "").toLowerCase().trim();
  if (!q) return [];
  const all = [
    ...STATUS_GUIDE.wr.map(s => ({ ...s, type: "Request" })),
    ...STATUS_GUIDE.wo.map(s => ({ ...s, type: "Work Order" })),
  ];
  return all.filter(s => s.party.toLowerCase().includes(q));
}

// Short summary string for system prompt context
function getSystemPromptSummary() {
  const parties = new Set();
  STATUS_GUIDE.wr.forEach(s => parties.add(s.party));
  STATUS_GUIDE.wo.forEach(s => parties.add(s.party));
  return "Responsible parties in the workflow: " + [...parties].join("; ");
}

module.exports = { STATUS_GUIDE, STATUS_INFO, getStatusInfo, listStatusesByParty, getSystemPromptSummary };
