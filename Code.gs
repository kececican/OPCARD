function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ Operasyon Paneli')
      .addItem('Dashboard\'u Aç', 'showDashboard')
      .addToUi();
}

function showDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(1100)
      .setHeight(800)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, 'Operasyon Kontrol Paneli');
}

// ─── Feedback Widget ───────────────────────────────────────────────────────────
var FB_SHEET_ID = '1ZdakLmkO8s57T-WU1MGuyMFitzRQxO5XxhuIL7UoMwI';

function submitFeedback(payload) {
  var ss      = SpreadsheetApp.openById(FB_SHEET_ID);
  var tabName = payload.appName || 'Genel';
  var sheet   = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.getRange(1, 1, 1, 9).setValues([[
      'FeedbackID','Email','Feedback_Type','Priority','Message',
      'CreatedAt','Comments','Status','Standardization Y/N'
    ]]).setBackground('#4A6CF7').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 120); sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 320); sheet.setColumnWidth(6, 160);
    sheet.setColumnWidth(7, 200);
  }

  var id    = Utilities.getUuid().substring(0, 8).toUpperCase();
  var now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var email = Session.getActiveUser().getEmail();

  sheet.appendRow([id, email,
    payload.feedbackType || '', payload.priority || '',
    payload.message || '', now, '', 'Open', '']);

  return { success: true, id: id };
}

// ─── Giriş/Çıkış Loglama ──────────────────────────────────────────────────────
var PROJECT_NAME = 'OPCARD';
var LOG_SHEET    = PROJECT_NAME + '_GirisLoglari';
var TIMEOUT_MIN  = 10;

function doGet() {
  var session  = createSession();
  var template = HtmlService.createTemplateFromFile('Index');
  template.sessionId  = session.sessionId;
  template.userEmail  = session.email;
  template.timeoutMin = TIMEOUT_MIN;
  return template.evaluate()
      .setTitle('Operasyon Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createSession() {
  var sheet     = _getOrCreateLogSheet();
  var sessionId = Utilities.getUuid();
  var email     = Session.getActiveUser().getEmail() || 'Anonim';
  sheet.appendRow([sessionId, email, new Date(), '', '', 'Açık']);
  return { sessionId: sessionId, email: email };
}

function logExit(sessionId, exitTimestamp, durationSec) {
  var sheet = _getOrCreateLogSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === sessionId && data[i][5] === 'Açık') {
      sheet.getRange(i + 1, 4).setValue(new Date(exitTimestamp));
      sheet.getRange(i + 1, 5).setValue(Math.round(durationSec / 60 * 10) / 10);
      sheet.getRange(i + 1, 6).setValue('Tamamlandı');
      break;
    }
  }
}

function _getOrCreateLogSheet() {
  var ss    = SpreadsheetApp.openById(FB_SHEET_ID);
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(['Oturum ID', 'E-posta', 'Giriş Zamanı', 'Çıkış Zamanı', 'Süre (dk)', 'Durum']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

// ─── Versiyon Yönetimi ────────────────────────────────────────────────────────

function getAppVersion() {
  var props = PropertiesService.getScriptProperties();
  return {
    version:   props.getProperty('APP_VERSION')    || '1.0',
    buildDate: props.getProperty('APP_BUILD_DATE') || '',
    isBeta:    props.getProperty('APP_IS_BETA')    !== 'false'
  };
}

function bumpVersion(type) {
  var props   = PropertiesService.getScriptProperties();
  var current = props.getProperty('APP_VERSION') || '1.0';
  var parts   = current.split('.').map(Number);
  if (type === 'major') { parts[0]++; parts[1] = 0; }
  else                  { parts[1]++; }
  var next      = parts[0] + '.' + parts[1];
  var buildDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  props.setProperty('APP_VERSION',    next);
  props.setProperty('APP_BUILD_DATE', buildDate);
  Logger.log('Versiyon: ' + next + ' (' + buildDate + ')');
  return { version: next, buildDate: buildDate };
}

function setBeta(flag) {
  PropertiesService.getScriptProperties().setProperty('APP_IS_BETA', flag ? 'true' : 'false');
  Logger.log('Beta modu: ' + (flag ? 'AÇIK' : 'KAPALI'));
}

function _kurulumYap(bumpType) {
  bumpVersion(bumpType);
}

function kurulumTrigger() { _kurulumYap('minor'); }

function majorDeploy()    { _kurulumYap('major'); }

function getOperationData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0]; 
  const data = sheet.getDataRange().getValues();
  
  return {
    mamulAdi: data[3] ? data[3][0] : "",
    referans: data[3] ? data[3][1] : "",
    opNo: data[5] ? data[5][0] : "",
    opAdi: data[5] ? data[5][1] : "",
    hatAdi: data[5] ? data[5][2] : "",
    makineNo: data[5] ? data[5][3] : "",
    parametre: data[5] ? data[5][4] : "",
    adimlar: data[9] ? [data[9][1], data[13][0], data[14][0], data[16][0], data[19][0]] : []
  };
}

function saveAndExportExcel(formData) {
  try {
    let imageUrl = "";
    if (formData.imageBase64) {
      const blob = dataURItoBlob(formData.imageBase64, formData.imageName);
      const folder = DriveApp.getRootFolder();
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = file.getUrl();
    }

    const currentSs = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = currentSs.getSheets()[0];
    const tempSs = SpreadsheetApp.create("OP_" + formData.referans);
    const newSheet = templateSheet.copyTo(tempSs);
    newSheet.setName(formData.referans);
    
    if (tempSs.getSheetByName("Sayfa1")) tempSs.deleteSheet(tempSs.getSheetByName("Sayfa1"));
    if (tempSs.getSheetByName("Sheet1")) tempSs.deleteSheet(tempSs.getSheetByName("Sheet1"));

    newSheet.getRange("A4").setValue(formData.mamulAdi);
    newSheet.getRange("B4").setValue(formData.referans);
    newSheet.getRange("A6").setValue(formData.opNo);
    newSheet.getRange("B6").setValue(formData.opAdi);
    newSheet.getRange("C6").setValue(formData.hatAdi);
    newSheet.getRange("D6").setValue(formData.makineNo);
    newSheet.getRange("E6").setValue(formData.parametre);
    
    formData.adimlar.forEach((adim, i) => {
      let cells = ["B10", "A14", "A15", "A17", "A20"];
      if(adim) newSheet.getRange(cells[i]).setValue(adim);
    });
    
    if(imageUrl) newSheet.getRange("D10").setFormula(`=IMAGE("${imageUrl}")`);

    const downloadUrl = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
    return { success: true, downloadUrl: downloadUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function dataURItoBlob(dataURI, fileName) {
  const splitData = dataURI.split(',');
  const type = splitData[0].split(':')[1].split(';')[0];
  const byteCharacters = Utilities.base64Decode(splitData[1]);
  return Utilities.newBlob(byteCharacters, type, fileName);
}
