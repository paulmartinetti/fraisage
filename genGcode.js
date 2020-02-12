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
  
  // get commun
  var sheet2 = sheetsA[1];
  var commun = sheet2.getName();
  
  /*** Creer programme principal (pp) ***/
  // 1. get 10 general vars from atelier (add # as we go)
  var dt = {};
  dt.str = "Dia_Tool";
  dt.val = ss.getRange("B3").getValue();
  
  var np = {};
  np.str = "Nb_Pass";
  np.val = ss.getRange("B5").getValue();
  
  var dc = {};
  dc.str = "D_Cut";
  dc.val = ss.getRange("B6").getValue();
  
  var c = {};
  c.str = "Clearance";
  c.val = ss.getRange("B8").getValue();
  
  // value with S
  var ssp = {};
  ssp.str = "Spindle_Speed";
  ssp.val = ss.getRange("B10").getValue();
  
  // value with F
  var rs = {};
  rs.str = "Rapid_Speed";
  rs.val = ss.getRange("B12").getValue();
  
  // value with F
  var ls = {};
  ls.str = "Low_Speed";
  ls.val = ss.getRange("B14").getValue();
  
  // value with F
  var ps = {};
  ps.str = "Plunge_Speed";
  ps.val = ss.getRange("B15").getValue();
  
  // drop-down list 1 - in, on, out
  var tm = {};
  tm.str = "Tool_Movement";
  tm.val = ss.getRange("B17").getValue();
  
  // drop-down list 2 - clockwise(col A) or counterCW(col B)
  var cd = {};
  cd.str = "Cutting_Direction";
  cd.val = ss.getRange("B18").getValue() == "clockwise" ? "A" : "B";
  //Logger.log(cd.val); - A
  
  // for substitution loop
  var subsA = [dt,np,dc,c,ssp,rs,ls,ps,tm,cd];
  
  // 11th var
  var mail = {};
  mail.str = "Email";
  mail.val = ss.getRange("B20").getValue();
  
  
  // 2. create coupeA array of all valid (non-zero) coupes to select Gcode for sp
  var coupeA = [];
  // get # de coupes
  var b1 = ss.getRange(atelierStr+'!B1').getValue();
  for (var i=2;i<=(b1*2);i+=2){
    var temp = "E"+i+":J"+i;
    var oneCoupe = sheet1.getRange(temp).getValues();
    // *** only push if non-zero, skip [0] which is coupe name
    var inclus = false;
    for (var j=1;j<oneCoupe[0].length;j++) {
      if (oneCoupe[0][j] > 0) {
        inclus = true;
        continue;
      }
    }
    if (inclus) coupeA.push(oneCoupe[0]);
  }
  // for looping
  var len = coupeA.length;
  //Logger.log(coupeA[0][0]); - Rectangle
  
  
  // 3. Create Commun strings (sheet2)
  // process Gcode arrays to strings with line breaks
  
  // Debut
  var debut = "";
  var gCodeA = sheet2.getRange("A2:A").getValues().filter(String);
  for (var k=0;k<gCodeA.length;k++){
    debut+=gCodeA[k]+n;
  }
  // Fin
  var fin = "";
  gCodeA = sheet2.getRange("B2:B").getValues().filter(String);
  // convert to string
  for (var k=0;k<gCodeA.length;k++){
    fin+=gCodeA[k]+n;
  }
  
  // 4. Iterate valid coupes to create middle (milieu)
  var milieu = "";
  
  // for each valid coupe
  for (i=0;i<len;i++) {
    // show cut number
    milieu+=n+"(Coupe "+(i+1)+" of "+len+")"+n;
    // spacer between cuts
    milieu+="(##################################)"+n;
    // get the corresponding sheet
    var sheetName = coupeA[i][0];
    // get array of values based on Cutting Direction (cd.val)
    gCodeA = ss.getRange(sheetName+"!"+cd.val+"2:"+cd.val).getValues().filter(String);
    //convert to string
    for (k=0;k<gCodeA.length;k++){
      milieu+=gCodeA[k]+n;
    }
    // substitute vars by looping once through row var names
    var parCoupeStrA = ["CutName","X","Y","Dia","Lg","Ht"];
    // start with X value
    for (j=1;j<parCoupeStrA.length;j++) {
        var pcStr = "#"+parCoupeStrA[j];
        //Logger.log(pcStr);
        var valStr = coupeA[i][j].toString();
        //Logger.log(valStr);
        var tRow = milieu.replace(pcStr, valStr);
        milieu = tRow
    }    
  }
  /*** HERE ***/
  Logger.log(milieu);
  // program principal
  var ppBeforeSubs = debut+milieu+fin;
  // make substitutions (not yet doing pecks or tool movement)
  var pp = "";
  // display
  //SpreadsheetApp.getUi().alert(ppBeforeSubs);
}

function majLatelier() {

  // get all current sheets to get coupes
  var sheetsA = ss.getSheets();
  
  // get atelier sheet name to program l'atelier
  var atelierStr = sheetsA[0].getName();
  
  // update the pulldowns based on all sheets, requires array
  var typeA = [];
  // note - skipping sheetsA[0, 1] l'atelier et commun
  for (var i=2;i<sheetsA.length;i++) {
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
