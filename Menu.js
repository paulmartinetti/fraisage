function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('gCode')
      .addItem('m-a-j Types', 'misajourTypes')
      .addItem('m-a-j Coupes','misajourCoupes')
      .addSeparator()
      .addItem('get G Code','chercheGcode')
      .addToUi();
}

function misajourTypes() {
  
  // sauvegarde Gcode sur Drive
  //DriveApp.createFile("LeDernierGcode.nc", getGcode(), MimeType.PLAIN_TEXT);
  
  // appelle-le
  //var file = DriveApp.getFilesByName("LeDernierGcode.nc");
  
  // creer brouillon avec le Gcode PJ
  /*  if (file.hasNext()) {
    GmailApp.createDraft('metal-fab@outlook.fr', 'Votre Gcode - Perceuse', 'Votre Gcode PJ automatic', {
      attachments: [file.next().getAs(MimeType.PLAIN_TEXT)],
      name: 'automated emailer script'
      });
   }*/
  
  majTypes();
}

function misajourCoupes() {
  SpreadsheetApp.getUi().alert("maj coupes");
}

function chercheGcode() {
  SpreadsheetApp.getUi().alert("get G Code");
}