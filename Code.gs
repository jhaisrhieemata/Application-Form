/**
 * Google Apps Script: Populate Google Doc template with form/sheet data,
 * save filled Doc, embed uploaded files, and export as PDF with shareable links.
 *
 * Update TEMPLATE_ID and OUTPUT_FOLDER_ID with your IDs.
 */

const TEMPLATE_ID = '1aLhrlA-rOhdbrPC3GeL7dR36AtScyozeuxO2iRMKAvg';  // Your template doc ID
const OUTPUT_FOLDER_ID = '1kZ28t0o7aKnG2abc6rqKWab_Pzx8t8CW';          // Folder to save filled docs
const RESPONSES_SHEET_NAME = 'Form Responses 1';                       // Sheet name

/**
 * Mapping: sheet header -> placeholder tag in doc
 */
const SHEET_HEADER_TO_TAG = {
  "branch_name": "{{branch_name}}",
  "landline": "{{landline}}",
  "address": "{{address}}",
  "mobile": "{{mobile}}",
  "date": "{{date}}",
  "client_name": "{{client_name}}",
  "nickname": "{{nickname}}",
  "contact_no": "{{contact_no}}",
  "client_address": "{{client_address}}",
  "birthplace": "{{birthplace}}",
  "birthdate": "{{birthdate}}",
  "age": "{{age}}",
  "civil_status": "{{civil_status}}",
  "dependents": "{{dependents}}",
  "religion": "{{religion}}",
  "valid_id": "{{valid_id}}",
  "fb_account": "{{fb_account}}",
  "height": "{{height}}",
  "weight": "{{weight}}",
  "citizenship": "{{citizenship}}",
  "acr_no": "{{acr_no}}",
  "issued_at": "{{issued_at}}",
  "resided_years": "{{resided_years}}",
  "resided_months": "{{resided_months}}",
  "Household Status": "{{Household Status}}", // custom checkboxes
  "proof_of_billing": "{{proof_of_billing}}",
  "landlord_name": "{{landlord_name}}",
  "previous_address": "{{previous_address}}",
  "father_name": "{{father_name}}",
  "father_occupation": "{{father_occupation}}",
  "mother_name": "{{mother_name}}",
  "mother_occupation": "{{mother_occupation}}",
  "family_address": "{{family_address}}",
  "spouse_name": "{{spouse_name}}",
  "spouse_contact": "{{spouse_contact}}",
  "spouse_occupation": "{{spouse_occupation}}",
  "spouse_employment_status": "{{spouse_employment_status}}",
  "spouse_rate": "{{spouse_rate}}",
  "spouse_company": "{{spouse_company}}",
  "spouse_company_contact": "{{spouse_company_contact}}",
  "spouse_company_address": "{{spouse_company_address}}",
  "position": "{{position}}",
  "monthly_salary": "{{monthly_salary}}",
  "company": "{{company}}",
  "company_contact": "{{company_contact}}",
  "company_address": "{{company_address}}",
  "years_employed": "{{years_employed}}",
  "agency_name": "{{agency_name}}",
  "agency_contact": "{{agency_contact}}",
  "agency_address": "{{agency_address}}",
  "other_income": "{{other_income}}",
  "other_income_amount": "{{other_income_amount}}",
  "other_income_address": "{{other_income_address}}",
  "note": "{{note}}",
  "comaker_name": "{{comaker_name}}",
  "comaker_contact": "{{comaker_contact}}",
  "comaker_relationship": "{{comaker_relationship}}",
  "comaker_nickname": "{{comaker_nickname}}",
  "comaker_age": "{{comaker_age}}",
  "comaker_address": "{{comaker_address}}",
  "comaker_occupation": "{{comaker_occupation}}",
  "comaker_salary": "{{comaker_salary}}",
  "comaker_company": "{{comaker_company}}",
  "comaker_years_employed": "{{comaker_years_employed}}",
  "comaker_company_address": "{{comaker_company_address}}",
  "comaker_company_contact": "{{comaker_company_contact}}",
  "model": "{{model}}",
  "color": "{{color}}",
  "downpayment": "{{downpayment}}",
  "terms": "{{terms}}",
  "driver_license": "{{driver_license}}", // file upload
  "sketch": "{{sketch}}"                  // file upload
};


/**
 * Generate from the latest row (last form response)
 */
function generateFromLatestResponse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + RESPONSES_SHEET_NAME + '" not found.');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('No response rows found in sheet.');

  const headers = values[0].map(h => (h || '').toString().trim());
  const lastRow = values[values.length - 1];
  const rowObj = {};

  for (let i = 0; i < headers.length; i++) {
    const key = headers[i];
    if (!key) continue;
    rowObj[key] = lastRow[i] !== undefined ? String(lastRow[i]) : '';
  }

  return createDocFromTemplate(rowObj);
}

/**
 * Create filled Doc + PDF
 */
function createDocFromTemplate(rowObj) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(),'dd-MM-yyyy');
  const applicantName = (rowObj.client_name || 'applicant').replace(/[^a-zA-Z0-9 _-]/g, '').substring(0, 40);
  const newDocName = `${applicantName}_${timestamp}`;

  const newFile = templateFile.makeCopy(newDocName, DriveApp.getFolderById(OUTPUT_FOLDER_ID));
  const newDocId = newFile.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // Fill placeholders
  for (const header in SHEET_HEADER_TO_TAG) {
    const placeholder = SHEET_HEADER_TO_TAG[header];
    let value = rowObj[header] || '';

    // Date formatting
    if (/date/i.test(header) && value) {
      value = toDateFmt(value);
    }

    // Household Status checkboxes
    if (header === "Household Status") {
      let ownedBox = (value.toLowerCase() === "owned") ? "☑" : "☐";
      let rentingBox = (value.toLowerCase() === "renting") ? "☑" : "☐";
      body.replaceText("{{owned_box}}", ownedBox);
      body.replaceText("{{renting_box}}", rentingBox);
      continue;
    }

    if (header === "driver_license" || header === "sketch") {
      insertFileAtPlaceholder(body, placeholder, value);
    } else {
      body.replaceText(escapeForRegExp(placeholder), value);
    }
  }

  // Apply font formatting
  formatDocumentFonts(body);

  doc.saveAndClose();

  // Export PDF
  const pdfBlob = DriveApp.getFileById(newDocId).getAs(MimeType.PDF).setName(newDocName + '.pdf');
  const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  const pdfFile = folder.createFile(pdfBlob);

  // Generate shareable links
  const links = generateShareableLinks(newDocId, pdfFile.getId());

  Logger.log("Doc Link: " + links.docUrl);
  Logger.log("PDF Link: " + links.pdfUrl);
  Logger.log("Shareable Doc Link: " + links.shareDocUrl);
  Logger.log("Shareable PDF Link: " + links.sharePdfUrl);

  return links;
}
function formatDocumentFonts(body) {
  const paras = body.getParagraphs();
  paras.forEach(p => {
    const text = p.editAsText();
    if (!text) return;

    // Force all text into fixed font
    text.setFontFamily("Times New Roman");
    text.setFontSize(11);
  });
}

/**
 * Insert uploaded file into the document
 */
function insertFileAtPlaceholder(body, placeholder, fileUrl) {
  if (!fileUrl) return;

  try {
    const fileIdMatch = fileUrl.match(/[-\w]{25,}/);
    if (!fileIdMatch) return;
    const fileId = fileIdMatch[0];
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();

    const found = body.findText(escapeForRegExp(placeholder));
    if (found) {
      const el = found.getElement();
      if (blob.getContentType().startsWith("image/")) {
        el.asText().replaceText(escapeForRegExp(placeholder), "");
        const image = el.getParent().asParagraph().insertInlineImage(0, blob);

        // 🔹 Scale proportionally
        const maxWidth = 720;
        const maxHeight = 480;
        let newWidth = image.getWidth();
        let newHeight = image.getHeight();

        if (newWidth > maxWidth) {
          const scale = maxWidth / newWidth;
          newWidth = newWidth * scale;
          newHeight = newHeight * scale;
        }
        if (newHeight > maxHeight) {
          const scale = maxHeight / newHeight;
          newWidth = newWidth * scale;
          newHeight = newHeight * scale;
        }

        image.setWidth(newWidth);
        image.setHeight(newHeight);
      } else {
        el.asText().replaceText(escapeForRegExp(placeholder), file.getUrl());
      }
    }
  } catch (err) {
    Logger.log("Error inserting file for " + placeholder + ": " + err);
  }
}

/**
 * Make shareable links (Doc + PDF)
 */
function generateShareableLinks(docId, pdfId) {
  const docFile = DriveApp.getFileById(docId);
  const pdfFile = DriveApp.getFileById(pdfId);

  docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    docUrl: `https://docs.google.com/document/d/${docId}/edit`,
    pdfUrl: pdfFile.getUrl(),
    shareDocUrl: docFile.getUrl(),
    sharePdfUrl: pdfFile.getUrl()
  };
}

/**
 * Escape regex chars in placeholder
 */
function escapeForRegExp(str) {
  return str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

/**
 * Add custom menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Template Tools')
    .addItem('Generate from latest response', 'generateFromLatestResponse')
    .addToUi();
}

/**
 * Trigger for onFormSubmit
 */
function onFormSubmitTrigger(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => (h || '').toString().trim());
  const values = e.values || [];
  const rowObj = {};
  for (let i = 0; i < headers.length; i++) {
    const key = headers[i];
    if (!key) continue;
    rowObj[key] = values[i] !== undefined ? String(values[i]) : '';
  }
  createDocFromTemplate(rowObj);
}

/**
 * Format datetimes to: DD-MM-YYYY
 */
function toDateFmt(dt_string) {
  var millis = Date.parse(dt_string);
  if (isNaN(millis)) return dt_string; // return raw if not valid date
  var date = new Date(millis);
  var day = ("0" + date.getDate()).slice(-2);
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var year = date.getFullYear();

  return `${day}-${month}-${year}`;
}

