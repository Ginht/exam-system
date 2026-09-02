const SPREADSHEET_ID = "1zOASIs7bnQR7QXbJeIHAvGICgUSizVjFAHgnqW6rhCk";

function getQuestions() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Questions');
  const data = sheet.getDataRange().getValues().slice(1);
  return data.map(r => ({id: r[0], question: r[1], answer: r[2]}));
}

function addQuestion(question, answer) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Questions');
  const lastRow = sheet.getLastRow();
  let newId = 1;
  if (lastRow > 1) {
    const lastId = sheet.getRange(lastRow, 1).getValue();
    newId = lastId + 1;
  }
  sheet.appendRow([newId, question, answer]);
  return {status:"ok"};
}

function deleteQuestion(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Questions');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i+1);
      return {status:"ok"};
    }
  }
  return {status:"notfound"};
}

function editQuestion(id, newQuestion, newAnswer) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Questions');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i+1, 2).setValue(newQuestion);
      sheet.getRange(i+1, 3).setValue(newAnswer);
      return {status:"ok"};
    }
  }
  return {status:"notfound"};
}

function doGet(e) {
  if (e.parameter.api === "questions") {
    return ContentService.createTextOutput(JSON.stringify(getQuestions()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({error:"Invalid API"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const action = e.parameter.action;
  let result = {};
  if (action === "add") {
    result = addQuestion(e.parameter.question, e.parameter.answer);
  } else if (action === "delete") {
    result = deleteQuestion(Number(e.parameter.id));
  } else if (action === "edit") {
    result = editQuestion(Number(e.parameter.id), e.parameter.question, e.parameter.answer);
  } else {
    result = {error:"Invalid action"};
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
