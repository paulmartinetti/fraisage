// global vars
var ss = SpreadsheetApp.getActive();

// commentaires
var clrCom = " (Move to clearance level)\n\n";
var n = "\n";

function majLatelier() {

  // get all current sheets
  var sheetsA = ss.getSheets();
  
  // capter Commande sheet name to program
  var commandeStr = sheetsA[0].getName();
  
  // update the pulldowns based on all sheets
  var typeA = [];
  // skipping sheetsA[0] which is Commande
  for (var i=1;i<sheetsA.length;i++) {
    // make obj and label pairs
    typeA.push(sheetsA[i].getName());
  }
  // set pulldown rule using array of coupe types
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(typeA);
  // get # of coupes
  var b1 = ss.getRange(commandeStr+'!B1').getValue();
  
  // m-a-j all pulldowns 
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
  Logger.log(typeA[1])
}

function getGcode() {

  // get all current sheets
  var sheetsA = ss.getSheets();
  // get Commande work area
  var sheet1 = sheetsA[0];
  var commandeStr = sheet1.getName();
  
  // build sous programme (sp)
  // 1. create coupeA array of all coupes to select valid Gcode for sp
  var coupeA = [];
  // get # de coupes
  var b1 = ss.getRange(commandeStr+'!B1').getValue();
  for (var i=2;i<=(b1*2);i+=2){
    var temp = "E"+i+":K"+i;
    var oneCoupe = sheet1.getRange(temp).getValues();
    // *** only push if non-Zero
    
    
    
    coupeA.push(oneCoupe[0]);
  }
   //Logger.log(coupeA[0][1]); = 1
  
  // 2. loop once through types, checking for each type in coupeA
  
  // skipping sheetsA[0] which is Commande
  for (i=1;i<sheetsA.length;i++) {
    // e.g. Circle
    var type = sheetsA[i].getName();
    // check coupeA for Circle
    for (var j=0;j<coupeA.length;j++) {
       // if Circle == Circle
       if (coupeA[j][0] == type) {
         // Assign an sp code O01+10+i
       }
    }
  }
}
  
