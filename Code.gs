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

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Operasyon Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

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
