// Google Apps Script 后端
// 部署为 Web App 时获取 URL，填入前端 API_URL

var SHEET_ID = 'YOUR_SHEET_ID_HERE'; // 替换为你的 Google Sheets ID

function doPost(e) {
  var data = JSON.parse(e.postData.contents);

  if (data.action === 'submit') {
    return handleSubmit(data);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = e.parameter.action;

  if (action === 'stats') {
    return handleStats(e.parameter.password);
  }
  if (action === 'export') {
    return handleExport(e.parameter.password);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSubmit(data) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('RawData');

  // 初始化表头（如果表为空）
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'timestamp', 'studentId',
      'phq9_1','phq9_2','phq9_3','phq9_4','phq9_5','phq9_6','phq9_7','phq9_8','phq9_9',
      'phq9_total', 'phq9_level',
      'gad7_1','gad7_2','gad7_3','gad7_4','gad7_5','gad7_6','gad7_7',
      'gad7_total', 'gad7_level'
    ]);
  }

  var phq9_total = data.phq9.reduce(function(a,b) { return a + b; }, 0);
  var gad7_total = data.gad7.reduce(function(a,b) { return a + b; }, 0);

  var phq9_level = getPHQ9Level(phq9_total);
  var gad7_level = getGAD7Level(gad7_total);

  var row = [
    new Date().toISOString(), data.studentId
  ].concat(data.phq9, phq9_total, phq9_level, data.gad7, gad7_total, gad7_level);

  sheet.appendRow(row);

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    phq9_total: phq9_total,
    phq9_level: phq9_level,
    gad7_total: gad7_total,
    gad7_level: gad7_level
  })).setMimeType(ContentService.MimeType.JSON);
}

function getPHQ9Level(score) {
  if (score <= 4) return '🟢 正常范围';
  if (score <= 9) return '🟡 轻度抑郁倾向';
  if (score <= 14) return '🟠 中度抑郁倾向';
  if (score <= 19) return '🔴 中重度抑郁倾向';
  return '🔴 重度抑郁倾向';
}

function getGAD7Level(score) {
  if (score <= 4) return '🟢 正常范围';
  if (score <= 9) return '🟡 轻度焦虑';
  if (score <= 14) return '🟠 中度焦虑';
  return '🔴 重度焦虑';
}

function handleStats(password) {
  if (password !== '123456') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('RawData');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({
      total: 0, levels: { good: 0, mild: 0, moderate: 0, severe: 0 },
      flagged: [], list: []
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var headers = data[0];
  var phq9TotalCol = headers.indexOf('phq9_total');
  var phq9LevelCol = headers.indexOf('phq9_level');
  var gad7TotalCol = headers.indexOf('gad7_total');
  var gad7LevelCol = headers.indexOf('gad7_level');
  var studentIdCol = headers.indexOf('studentId');
  var phq9Q9Col = headers.indexOf('phq9_9');

  var levels = { good: 0, mild: 0, moderate: 0, severe: 0 };
  var flagged = [];
  var list = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var phq9 = row[phq9TotalCol];
    var gad7 = row[gad7TotalCol];
    var studentId = row[studentIdCol];

    // 取较严重的等级
    if (phq9 <= 4 && gad7 <= 4) levels.good++;
    else if (phq9 <= 9 && gad7 <= 9) levels.mild++;
    else if (phq9 <= 14 && gad7 <= 14) levels.moderate++;
    else levels.severe++;

    list.push({
      studentId: studentId,
      phq9: phq9,
      phq9_level: row[phq9LevelCol],
      gad7: gad7,
      gad7_level: row[gad7LevelCol],
      danger: row[phq9Q9Col] >= 1
    });

    if (row[phq9Q9Col] >= 1) {
      flagged.push(studentId);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    total: data.length - 1,
    levels: levels,
    flagged: flagged,
    list: list
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleExport(password) {
  if (password !== '123456') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('RawData');
  var data = sheet.getDataRange().getValues();

  function csvEscape(val) {
    var s = String(val);
    if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  var csv = data.map(function(row) {
    return row.map(csvEscape).join(',');
  }).join('\n');

  return ContentService.createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV);
}
