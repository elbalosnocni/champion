// ĐIỀN GOOGLE SHEET ID CỦA BẠN VÀO GIỮA HAI DẤU NHÁY ĐƠN
const SHEET_ID = '1j24IRF6hR5iD-aSM_5HS9Cv5oDK4TMayVzQIaB_5mfY'; 

function getTargetSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
}

function doGet(e) {
  var sheet = getTargetSheet();
  var rows = sheet.getDataRange().getValues();
  var data = [];
  
  for (var i = 1; i < rows.length; i++) {
    data.push({
      id: rows[i][0],
      teamName: rows[i][1],
      captain: rows[i][2],
      players: JSON.parse(rows[i][3] || "[]"),
      status: rows[i][4],
      matches: JSON.parse(rows[i][5] || "[]")
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = getTargetSheet();
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "TeamName", "Captain", "Players JSON", "Status", "Matches JSON"]);
  }
  
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  
  if (action === "register") {
    sheet.appendRow([
      params.id,
      params.teamName,
      params.captain,
      JSON.stringify(params.players),
      "Chờ duyệt",
      JSON.stringify([])
    ]);
    return ContentService.createTextOutput("Success");
  } 
  
  if (action === "approve") {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() == params.id.toString()) {
        sheet.getRange(i + 1, 5).setValue("Đã duyệt");
        break;
      }
    }
    return ContentService.createTextOutput("Approved");
  }
  
  if (action === "updateMatch") {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1].toString() == params.teamName.toString()) {
        sheet.getRange(i + 1, 6).setValue(JSON.stringify(params.matches));
      }
    }
    return ContentService.createTextOutput("Match Updated");
  }
}
