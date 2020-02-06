// global vars
var ss = SpreadsheetApp.getActive();

// les types de coupe et leur gCode
var typeA = [];
var gcodeA = [];

// commentaires
var clrCom = " (Move to clearance level)\n\n";
var n = "\n";


function majTypes() {
  
  // get all current sheets
  var sheetsA = ss.getSheets();
  
  // supprime Commande sheet
  var commandeStr = sheetsA.splice(0,1)[0].getName();
  var sLen = sheetsA.length;
  
  // new vars from sheets into gcodeA
  // update the pulldown
  for (var i=0;i<sLen;i++) {
    // make obj and label pairs
    typeA.push(sheetsA[i].getName());
    gcodeA.push(sheetsA[i].getRange("A1:A").getValues().filter(String));
  }
  
  // set pulldown rule (array of types)
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(typeA);
  // get # de coupes
  var b1 = ss.getRange(commandeStr+'!B1').getValue();
  
  // maj pulldowns 
  for (var i=0;i<b1;i++) {
    // col E, on even rows 2+
    var temp = commandeStr+'!E'+((i*2)+2);
    ss.getRange(temp).setDataValidation(rule);
    
    temp = commandeStr+'!D'+((i*2)+2);
    // col D, add order
    ss.getRange(temp).setValue(i+1);
    
  }
}
