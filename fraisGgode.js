// global vars
var ss = SpreadsheetApp.getActive();

// les types de coupe et leur gCode
var typeA = [];
var gcodeA = [];

// commentaires
var clrCom = " (Move to clearance level)\n\n";
var n = "\n";

function majCoupes() {
  
  // get Commande work area
  var commande = ss.getSheets()[0].getName();
  
  // get # de coupes
  var b1 = ss.getRange(commande+'!B1').getValue();
  
  //Logger.log(b1);
}

function majTypes() {

  // get all current sheets
  var sheetsA = ss.getSheets();
  
  // separarer Commande sheet du reste
  var commandeStr = sheetsA.splice(0,1)[0].getName();
  
  // new vars from sheets into gcodeA
  // update the pulldown
  for (var i=0;i<sheetsA.length;i++) {
    // make obj and label pairs
    typeA.push(sheetsA[i].getName());
    gcodeA.push(sheetsA[i].getRange("A1:A").getValues().filter(String));
  }
  
  // set pulldown rule (array of types)
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(typeA);
  // get # de coupes
  var b1 = ss.getRange(commandeStr+'!B1').getValue();
  
  // m-a-j pulldowns 
  var temp = "";
  var dig = 0;
  for (var i=0;i<b1;i++) {
    // col E, on even rows 2+
    dig = ((i*2)+2);
    temp = commandeStr+'!E'+dig;
    ss.getRange(temp).setDataValidation(rule);
    
    // col D, add order
    temp = commandeStr+'!D'+dig;
    ss.getRange(temp).setValue(i+1);
    
    // colorize rows
    temp = commandeStr+'!D'+dig+':K'+dig;
    ss.getRange(temp).setBackground('#cccccc');
  }
}
