// =====================================================================
// APPROVAL WORKFLOW SERVICE
// =====================================================================

/**
 * Submits a new request from a Project to HQ.
 * 1. Generates a Request ID.
 * 2. Saves to Local Project Sheet (Approval_Requests_Local).
 * 3. Saves to HQ Sheet (Central_Approvals).
 * @param {string} projectKey - The key from Config (e.g., 'LIMB_CENTRE', 'DISPENSARY')
 * @param {Object} requestData - Object containing {Patient_Name, Type, Cost}
 */
function submitRequest(projectKey, requestData) {
  try {
    // 1. Generate a unique Request ID (REQ-Timestamp)
    const requestId = "REQ-" + new Date().getTime();

    // 2. Save to Local Project File first
    // Mapping data to match "Approval_Requests_Local" columns
    const localData = {
      'Request_ID': requestId,
      'Patient_ID': requestData.Patient_ID || 'N/A',
      'Patient_Name': requestData.Patient_Name,
      'Request_Type': requestData.Request_Type,
      'Cost_Estimate': requestData.Cost || '',
      'Priority': requestData.Priority || 'Normal',
      'Request_Notes': requestData.Request_Notes || '',
      'Status_Synced_From_HQ': 'Pending'
    };

    // Note: The tab name MUST match your sheet exactly: 'Approval_Requests_Local'
    addData(projectKey, 'Approval_Requests_Local', localData);

    // 3. Save to HQ File
    // Mapping data to match "Central_Approvals" columns
    const hqData = {
      'Request_ID': requestId,
      'Source_Project': projectKey,
      'Request_Date': new Date(), // Current Date
      'Request_Type': requestData.Request_Type,
      'Patient_Client_Name': requestData.Patient_Name,
      'Cost_Estimate': requestData.Cost || '',
      'Priority': requestData.Priority || 'Normal',
      'Request_Notes': requestData.Request_Notes || '',
      'Approval_Status': 'Pending',
      'Approver_Comments': ''
    };

    // Note: The tab name MUST match your sheet exactly: 'Central_Approvals'
    addData('HQ', 'Central_Approvals', hqData);

    return { status: "Success", id: requestId, message: "Request sent to HQ successfully." };

  } catch (e) {
    Logger.log("Error in submitRequest: " + e.toString());
    return { status: "Error", message: e.toString() };
  }
}

/**
 * Approves or Rejects a request (Run by Admin/HQ).
 * Updates HQ and syncs back to the Source Project.
 * @param {string} requestId - The ID to look for (e.g., REQ-123456)
 * @param {string} decision - "Approved" or "Rejected"
 * @param {string} comments - Reason for rejection or notes
 */
function processApproval(requestId, decision, comments) {
  try {
    const hqSS = getSpreadsheet('HQ');
    const hqSheet = hqSS.getSheetByName('Central_Approvals');
    const hqData = hqSheet.getDataRange().getValues();

    let sourceProject = "";
    let found = false;
    let rowIndex = 0;

    // 1. Find the row in HQ Sheet
    // We start at i=1 to skip headers
    for (let i = 1; i < hqData.length; i++) {
      // Assuming Request_ID is in Column 1 (Index 0)
      if (hqData[i][0] == requestId) {
        rowIndex = i + 1; // Convert 0-based index to 1-based row number
        sourceProject = hqData[i][1]; // Column 2 is Source_Project
        found = true;
        break;
      }
    }

    if (!found) return "Error: Request ID not found in HQ Database.";

    // 2. Update HQ Sheet
    // Column 8 is Approval_Status, Column 9 is Approver_Comments
    hqSheet.getRange(rowIndex, 8).setValue(decision);
    hqSheet.getRange(rowIndex, 9).setValue(comments);

    // 3. Sync to Source Project Sheet
    // We check if the source project exists in our Config
    if (sourceProject && sourceProject !== "") {
      try {
        const projSS = getSpreadsheet(sourceProject);
        const projSheet = projSS.getSheetByName('Approval_Requests_Local');

        if(projSheet) {
          const projData = projSheet.getDataRange().getValues();
          for (let j = 1; j < projData.length; j++) {
            // Assuming Request_ID is in Column 1 (Index 0)
            if (projData[j][0] == requestId) {
               // Update Status in Local Sheet (Column 7: Status_Synced_From_HQ)
               projSheet.getRange(j + 1, 7).setValue(decision);
               break;
            }
          }
        }
      } catch (err) {
        // If we can't open the project sheet, we log it but don't fail the HQ update
        Logger.log("Could not sync back to project sheet: " + err.toString());
        return "Updated HQ, but failed to sync to Project Sheet (Check Sheet Name).";
      }
    }

    return "Success: Request processed and synced to project.";

  } catch (e) {
    Logger.log("Error in processApproval: " + e.toString());
    return "Error: " + e.toString();
  }
}