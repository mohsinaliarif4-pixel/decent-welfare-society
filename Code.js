function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('DWS ERP Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPageContent(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return `<div class="p-10 text-center text-red-500 font-bold">Error 404: Page '${filename}' not found.</div>`;
  }
}

// Alias for older calls
function getPage(f) { return getPageContent(f); }

// --- AUTHENTICATION ---
function verifyUserLogin(email) {
  try {
    const cleanInputEmail = email.toString().trim().toLowerCase();
    Logger.log("Attempting login for: " + cleanInputEmail);

    // 1. Connect to HQ Sheet
    const ss = SpreadsheetApp.openById("1C9i-zOXJ5piESoxjyPsru6lIKQepSRDPzNjgjo3UxuA");

    // 2. Access Tab by Name - "System_Users"
    let sheet = ss.getSheetByName("System_Users");

    if(!sheet) {
        // Fallback check
        sheet = ss.getSheetByName("Range_System_Users");
    }

    if(!sheet) {
        Logger.log("Critical Error: Tab 'System_Users' not found.");
        return { allowed: false, error: "System Configuration Error: User DB missing." };
    }

    // 3. Get Data using getDisplayValues for robust string matching
    const data = sheet.getDataRange().getDisplayValues();

    // Headers are in Row 1 (Index 0). Data starts Row 2 (Index 1).
    // Columns: Name(A), User_Email(B), Role(C), Permissions(D), Status(E)
    // Indexes: 0, 1, 2, 3, 4

    for(let i = 1; i < data.length; i++) {
        const row = data[i];

        // Skip empty rows
        if(row.join("").trim() === "") continue;

        const dbEmail = row[1].toString().trim().toLowerCase();
        const status = row[4].toString().trim();

        if (dbEmail === cleanInputEmail) {
            Logger.log("Match Found at row " + (i+1));
            if(status === 'Active') {
                return {
                    allowed: true,
                    name: row[0],
                    role: row[2],
                    permissions: row[3]
                };
            } else {
                return { allowed: false, error: "Account is Suspended" };
            }
        }
    }

    Logger.log("Access Denied: Email not found.");
    return { allowed: false, error: "Email not found in database." };

  } catch (e) {
    Logger.log("Login Exception: " + e.toString());
    return { allowed: false, error: "System Error: " + e.toString() };
  }
}

function getUserRole() {
    // Legacy function support - defaults to guest until login
    return { role: 'Guest', department: 'HQ', name: 'Guest User' };
}