/**
 * Google Apps Script backend for Takht-e Jamshid Driving School
 *
 * 1) script.google.com -> New project
 * 2) این کد را در Code.gs قرار دهید.
 * 3) Deploy -> New deployment -> Web app
 * 4) Execute as: Me
 * 5) Who has access: Anyone
 * 6) آدرس Web App را در index.html داخل API_URL قرار دهید.
 *
 * داده‌ها در یک Google Sheet داخل Google Drive شما ذخیره می‌شوند.
 * اگر SHEET_ID خالی باشد، یک فایل جدید ساخته می‌شود.
 */

const SHEET_ID = ""; // اختیاری: ID یک Google Sheet موجود
const SHEET_NAME = "ثبت‌نام‌ها";

function getSheet_() {
  let ss;
  if (SHEET_ID) {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } else {
    const props = PropertiesService.getScriptProperties();
    const savedId = props.getProperty("SHEET_ID");
    if (savedId) {
      ss = SpreadsheetApp.openById(savedId);
    } else {
      ss = SpreadsheetApp.create("ثبت نام آموزشگاه رانندگی تخت جمشید");
      props.setProperty("SHEET_ID", ss.getId());
    }
  }

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "تاریخ",
      "نوع درخواست",
      "نام و نام خانوادگی",
      "شماره موبایل",
      "کد ملی",
      "سن",
      "ترجیح مربی",
      "زمان ترجیحی",
      "توضیحات"
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet() {
  return json_({ok:true, message:"Takht-e Jamshid registration API is active."});
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const sh = getSheet_();

    const type = data.type === "coach_request" ? "درخواست انتخاب مربی" : "ثبت‌نام آموزشگاه";

    sh.appendRow([
      new Date(),
      type,
      data.name || "",
      data.phone || "",
      data.national_id || "",
      data.age || "",
      data.coach_preference || "",
      data.preferred_time || "",
      data.notes || ""
    ]);

    return json_({ok:true});
  } catch (err) {
    return json_({ok:false, error:String(err)});
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
