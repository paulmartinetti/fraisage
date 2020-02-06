function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('G code')
      .addItem("m-a-j l'atelier", 'misaJourLatelier')
      .addSeparator()
      .addItem('get G code','chercheGcode')
      .addToUi();
}

function misaJourLatelier() {
  
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
  
  majLatelier();
}

function chercheGcode() {
  SpreadsheetApp.getUi().alert("get G code");
}