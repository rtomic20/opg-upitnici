// Google Apps Script — OPG Priča Upitnik
// Sheet ID: 1h_F1Qz7Mm-vhXd0RvFrs8bcJabjKf3NTqCiHjeiFhuA
//
// Upute za deploy:
// 1. Otvori: https://script.google.com
// 2. Stvori novi projekt (+ New project)
// 3. Zalijepi ovaj kod (zamijeni sav postojeći)
// 4. Klikni Deploy → New deployment
// 5. Type: Web app
// 6. Execute as: Me
// 7. Who has access: Anyone
// 8. Klikni Deploy i kopiraj URL
// 9. Zalijepi URL u prica.html u varijablu SCRIPT_URL

const SHEET_ID = "1h_F1Qz7Mm-vhXd0RvFrs8bcJabjKf3NTqCiHjeiFhuA";

const HEADERS = [
  "Timestamp",
  "OPG naziv",
  "Kontakt ime",
  "Email",
  "Telefon",
  "1. Nastanak OPG-a",
  "2. Motivacija",
  "3. Tipičan dan",
  "4. Posebnost proizvoda",
  "5. Najpoznatiji proizvod",
  "6. OPG u jednoj rečenici",
  "7. Posebna uspomena",
  "8. Fotografije",
  "8. Fotografije - napomena",
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("Priče OPG-ova");

    // Stvori sheet ako ne postoji i dodaj zaglavlja
    if (!sheet) {
      sheet = ss.insertSheet("Priče OPG-ova");
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const row = [
      new Date().toLocaleString("hr-HR"),
      data.opg_naziv      || "",
      data.kontakt_ime    || "",
      data.kontakt_email  || "",
      data.kontakt_telefon|| "",
      data.q1             || "",
      data.q2             || "",
      data.q3             || "",
      data.q4             || "",
      data.q5             || "",
      data.q6             || "",
      data.q7             || "",
      data.q8             || "",
      data.q8_napomena    || "",
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test funkcija — pokreni ručno da provjeriš pristup sheetu
function testWrite() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log("Sheet name: " + ss.getName());
}
