// ==========================================
// CONFIGURATION & IDS
// ==========================================

const SHEET_IDS = {
  // 01_DWS_Main_HQ
  HQ: "1C9i-zOXJ5piESoxjyPsru6lIKQepSRDPzNjgjo3UxuA",

  // 02_DWS_Limb_Centre
  LIMB_CENTRE: "1Ap057Z2Jh5-xuaDsEXQApZz87hgVDtPrTXyBmc41Jxw",

  // 03_DWS_Dispensary
  DISPENSARY: "18Pegr9qRDaCE5PEnS-1KQLxrIbSaziRLt-Dh1H8VTsQ",

  // 04_DWS_Ration
  RATION: "1sIZ6IXMKd74v2G0iEJKIK1jqLVmC4_8RqLIDifK67r0",

  // 05_DWS_Health_Camps
  CAMPS: "1sX74zsgFqb7s5bQGWnlFkcgIgcNeB2Tv57a5mm2haM8"
};

const DRIVE_FOLDERS = {
  // Folder for Patient Images and Transaction Evidence
  PATIENTS: "1R7Brrb1gfQe_ihayIeC_STJw8GeSC29X"
};

function getSpreadsheet(key) {
  if (!SHEET_IDS[key]) throw new Error("Invalid Sheet Key: " + key);
  return SpreadsheetApp.openById(SHEET_IDS[key]);
}