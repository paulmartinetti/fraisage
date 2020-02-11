// global vars
var ss = SpreadsheetApp.getActive();

// commentaires
var n = "\n";

function getGcode() {

  // get all current sheets
  var sheetsA = ss.getSheets();
  // get atelier work area
  var sheet1 = sheetsA[0];
  var atelierStr = sheet1.getName();
  
  /*** build sous programme (sp) ***/
  // 1. create coupeA array of all valid coupes to select Gcode for sp
  var coupeA = [];
  // get # de coupes
  var b1 = ss.getRange(atelierStr+'!B1').getValue();
  for (var i=2;i<=(b1*2);i+=2){
    var temp = "E"+i+":K"+i;
    var oneCoupe = sheet1.getRange(temp).getValues();
    // *** only push if non-Zero, skip [0] which is coupe name
    var inclus = false;
    for (var j=1;j<oneCoupe[0].length;j++) {
      if (oneCoupe[0][j] > 0) {
        inclus = true;
        continue;
      }
    }
    if (inclus) coupeA.push(oneCoupe[0]);
  }
  
  // get all vars
  
}

function majLatelier() {

  // get all current sheets to get coupes
  var sheetsA = ss.getSheets();
  
  // get atelier sheet name to program l'atelier
  var atelierStr = sheetsA[0].getName();
  
  // update the pulldowns based on all sheets, requires array
  var typeA = [];
  // note - skipping sheetsA[0] which is atelier
  for (var i=1;i<sheetsA.length;i++) {
    // make obj and label pairs
    typeA.push(sheetsA[i].getName());
  }
  // set pulldown rule using array of coupe types
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(typeA);
  // get # of coupes
  var b1 = ss.getRange(atelierStr+'!B1').getValue();
  
  // m-a-j all pulldowns 
  var temp = "";
  var dig = 0;
  for (var i=0;i<b1;i++) {
    // col E, on even rows 2+
    dig = ((i*2)+2);
    temp = atelierStr+'!E'+dig;
    ss.getRange(temp).setDataValidation(rule);
    
    // col D, add order
    temp = atelierStr+'!D'+dig;
    ss.getRange(temp).setValue(i+1);
    
    // colorize rows
    temp = atelierStr+'!D'+dig+':J'+dig;
    ss.getRange(temp).setBackground('#cccccc');
  }
}
