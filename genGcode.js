// global vars
var ss = SpreadsheetApp.getActive();

// vars reference
// 1. loops - k, i, j
// 2. grid vars - dt, tm, cd, mail, subsA

// commentaires
var n = "\n";

function getGcode() {

  // get all current sheets
  var sheetsA = ss.getSheets();

  // get atelier work area
  var sheet1 = sheetsA[0];

  // get commun
  var sheet2 = sheetsA[1];
  var commun = sheet2.getName();

  /*** Creer programme principal (pp) ***/
  // 1. get 10 general vars from atelier (add # as we go)
  var dt = {};
  dt.str = "Dia_Tool";
  dt.val = sheet1.getRange("B3").getValue();

  // for substitution loop
  var subsA = [
    { str: dt.str, val: dt.val },
    { str: "Nb_Pass", val: sheet1.getRange("B5").getValue() },
    { str: "D_Cut", val: sheet1.getRange("B6").getValue() },
    { str: "Clearance", val: sheet1.getRange("B8").getValue() },
    { str: "Spindle_Speed", val: sheet1.getRange("B10").getValue() },
    { str: "Rapid_Speed", val: sheet1.getRange("B12").getValue() },
    { str: "Low_Speed", val: sheet1.getRange("B14").getValue() },
    { str: "Plunge_Speed", val: sheet1.getRange("B15").getValue() }
  ];

  // drop-down list 1 - in, on, out, calculate delta
  var tm = {};
  tm.str = "Tool_Movement";
  tm.val = sheet1.getRange("B17").getValue();
  tm.delta = 0;
  if (tm.val == "tool inside form") {
    tm.delta = -1 * Math.round((dt.val / 2), 1);
  } else if (tm.val == "tool outside form") {
    tm.delta = Math.round(dt.val / 2, 1);
  }

  // drop-down list 2 - clockwise(col A) or counterCW(col B)
  var cd = {};
  cd.str = "Cutting_Direction";
  cd.val = sheet1.getRange("B18").getValue() == "horaire" ? "A" : "B";
  //Logger.log(cd.val); - A

  // email address to send G codde
  var mail = {};
  mail.str = "Email";
  mail.val = sheet1.getRange("B20").getValue();

  // 2. Create Commun strings (sheet2)
  // process Gcode arrays to strings with line breaks

  // Debut
  var debut = "";
  // get sheet 2 (Commun) column A
  var gCodeA = sheet2.getRange("A2:A").getValues().filter(String);
  // convert array to string
  for (var k = 0; k < gCodeA.length; k++) {
    var rep = subMe(gCodeA[k].toString(), subsA) + n;
    debut += rep;
  }

  // Fin
  var fin = n;
  gCodeA = sheet2.getRange("B2:B").getValues().filter(String);
  // convert to string
  for (k = 0; k < gCodeA.length; k++) {
    rep = subMe(gCodeA[k].toString(), subsA) + n;
    fin += rep;
  }

  // 3. create coupeA array of all valid (non-zero) coupes to select Gcode for sp
  var coupeA = [];
  // get # de coupes
  var b1 = sheet1.getRange("B1").getValue();
  // i is rowNum, not an array index!
  for (i = 2; i <= (b1 * 2); i += 2) {
    var temp = "E" + i + ":J" + i;
    var oneCoupe = sheet1.getRange(temp).getValues();
    // *** only push if non-zero, skip [0] which is coupe name
    var inclus = false;
    for (var j = 1; j < oneCoupe[0].length; j++) {
      if (oneCoupe[0][j] > 0) {
        inclus = true;
        continue;
      }
    }
    if (inclus) {
      // cutsA to store labeled values for each coupe
      var cutsA = [
        { str: 'Type', val: "" },
        { str: 'X', val: 0 },
        { str: 'Y', val: 0 },
        { str: 'Dia', val: 0 },
        { str: 'Lg', val: 0 },
        { str: 'Ht', val: 0 }
      ];
      var cutLen = cutsA.length;
      // convert oneCoupe[0] into cutsA, labeling its values
      for (k = 0; k < cutLen; k++) {
        // for each coupe value
        var cval = oneCoupe[0][k];
        var cut = cutsA[k];
        // update Lg, Ht for tool movement using tm.delta before storing
        if (cut.str == "Lg" || cut.str == "Ht") {
          cut.val = cval + tm.delta;
          continue;
        }
        // store rest
        cut.val = cval;
      }
      // une coupe, X and Y maj, str/val pairs for substition, ready to use
      coupeA.push(cutsA);
    }
  }

  // 4. Iterate valid coupes to create middle (milieu)
  var milieu = "";
  var len = coupeA.length;
  // for each valid coupe
  for (i = 0; i < len; i++) {
    // show cut number
    milieu += n + "(Coupe " + (i + 1) + " of " + len + ")" + n;
    // spacer between cuts
    milieu += "(##################################)" + n;
    // get the corresponding sheet
    var sheetName = coupeA[i][0].val;
    // get array of values based on Cutting Direction (cd.val)
    gCodeA = ss.getRange(sheetName + "!" + cd.val + "2:" + cd.val).getValues().filter(String);
    //convert to string
    for (k = 0; k < gCodeA.length; k++) {
      // 1. subsA 
      rep = subMe(gCodeA[k].toString(), subsA);
      // 2. coupeA[i] (cutsA)
      rep= subMe(rep, coupeA[i]) + n;
      milieu += rep;
    }
  }
  // ****** Z !! - not yet doing
  // program principal
  var pp = debut + milieu + fin;
  // display
  SpreadsheetApp.getUi().alert(pp);
  // sauvegarde Gcode sur Drive
  //DriveApp.createFile("LeDernierGcode.nc", pp, MimeType.PLAIN_TEXT);
}

function subMe(lineTxt, subA) {
  for (var i = 0; i < subA.length; i++) {
    var tStr = "#" + subA[i].str;
    var tVal = subA[i].val.toString();
    var rep = lineTxt.replace(tStr, tVal);
    lineTxt = rep;
  }
  return lineTxt;
}

function majLatelier() {

  // get all current sheets to get coupes
  var sheetsA = ss.getSheets();

  // get atelier sheet name to program l'atelier
  var atelierStr = sheetsA[0].getName();

  // update the pulldowns based on all sheets, requires array
  var typeA = [];
  // note - skipping sheetsA[0, 1] l'atelier et commun
  for (var i = 2; i < sheetsA.length; i++) {
    // make obj and label pairs
    typeA.push(sheetsA[i].getName());
  }
  // set pulldown rule using array of coupe types
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(typeA);
  // get # of coupes
  var b1 = ss.getRange(atelierStr + '!B1').getValue();

  // m-a-j all pulldowns 
  var temp = "";
  var dig = 0;
  for (var i = 0; i < b1; i++) {
    // col E, on even rows 2+
    dig = ((i * 2) + 2);
    temp = atelierStr + '!E' + dig;
    ss.getRange(temp).setDataValidation(rule);

    // col D, add order
    temp = atelierStr + '!D' + dig;
    ss.getRange(temp).setValue(i + 1);

    // colorize rows
    temp = atelierStr + '!D' + dig + ':J' + dig;
    ss.getRange(temp).setBackground('#cccccc');
  }
}
