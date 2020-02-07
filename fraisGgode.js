// global vars
var ss = SpreadsheetApp.getActive();

// commentaires
var clrCom = " (Move to clearance level)\n\n";
var n = "\n";

function majLatelier() {

  // get all current sheets to get coupes
  var sheetsA = ss.getSheets();
  
  // get Commande sheet name to program l'atelier
  var commandeStr = sheetsA[0].getName();
  
  // update the pulldowns based on all sheets, requires array
  var typeA = [];
  // note - skipping sheetsA[0] which is Commande
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
}

function getGcode() {

  // get all current sheets
  var sheetsA = ss.getSheets();
  // get Commande work area
  var sheet1 = sheetsA[0];
  var commandeStr = sheet1.getName();
  
  /*** build sous programme (sp) ***/
  // 1. create coupeA array of all valid coupes to select Gcode for sp
  var coupeA = [];
  // get # de coupes
  var b1 = ss.getRange(commandeStr+'!B1').getValue();
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
  //Logger.log(coupeA);
  
  
  // 2. loop once through types, capturing Gcode sp for each type in coupeA
  var sp = "";
  var typeA = [];
  // skipping sheetsA[0] which is Commande
  for (i=1;i<sheetsA.length;i++) {
    // e.g. Circle, sheet 2 in sheetsA, the first type sheet
    var type = sheetsA[i].getName();
    // check first element of all arrays in coupeA for Circle
    var match = false;
    for (j=0;j<coupeA.length;j++) {
       // if Circle == Circle at least once
       if (coupeA[j][0] == type) {
         match = true;
         // good enough
         continue;
       }
    }
    // if match, add Gcode to sp
    if (match) {
    // Assign an sp code O01+10+i
         sp+="O12"+(10+i)+" (sp)"+n;
         // process Gcode arrays to strings with line breaks
         var gStr = "";
         var gCodeA = sheetsA[i].getRange("A1:A").getValues().filter(String);
         for (var k=0;k<gCodeA.length;k++){
           gStr+=gCodeA[k]+n;
         }
         // append to sp
         sp+=gStr;
         // spacer
         sp+=n;
         // save type obj for pp
         var obj = {};
         obj.type = type;
         obj.ind = i;
         typeA.push(obj);
    }
  }
  
  /*** build programme principal (pp) ***/
  // debut
  var pp = "G90 (Absolute positioning)"+n
    +"G1 X0 Y0 F1200"+n
    +"G1 Z0 F1200"+clrCom
    +"(commence programme principal)"+n;
    
  // 1. use coupeA (array of rows in atelier) for valid coupes
  // ignore[0] which is name
  for (i=0;i<coupeA.length;i++){
    // capture coupe type
    type =coupeA[i][0];
    // update coupe #
    pp+=n;
    pp+="(Coupe No. "+(i+1)+", "+type+")"+n;
    // update vars based on index
    for (j=1;j<coupeA[i].length;j++){
      pp+="#1"+(10+j)+"="+coupeA[i][j]+n;
    }
    // for each coupe, get matching sheet index from typeA for sp call
    for (k=0;k<typeA.length;k++){
      if (typeA[k].type == type) {
        pp+="M98 P12"+(10+typeA[k].ind)+" (appel sp)"+n;
        continue;
      }
    }
  }
  pp+="M30 (fin pp)";
  return pp+n+n+n+sp;
}
  
