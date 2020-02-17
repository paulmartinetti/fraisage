// global vars
var ss = SpreadsheetApp.getActive();

/**
 * vars reference
 * 1. loops - k, i, j
 * 2. grid vars - dt, tm, cd, mail, subsA, coupeA
 * 3. coupe = cut = Type, X, Y, Lg, Ht
 */

/**
 *   for V3
 * 1. Z - pecking
 * 2. Btn to show required params per type
 *  */

// commentaires
var n = "\n";

function getGcode() {

  // get all current sheets
  var sheetsA = ss.getSheets();

  // get atelier work area
  var atelier = sheetsA[0];

  // get commun
  var commun = sheetsA[1];

  /*** Creer programme principal (pp) ***/
  // 1. get vars that are the same for all cuts
  // Diameter of the tool
  var dt = {};
  dt.str = "Dia_Tool";
  dt.val = atelier.getRange("B3").getValue();

  // drop-down list 1 - in, on, out, calculate delta for Lg, Ht
  // e.g. if tool is inside, subtract half the tool's diameter
  // from the width and height
  var tm = {};
  tm.str = "Tool_Movement";
  tm.val = atelier.getRange("B17").getValue();
  tm.delta = 0;
  
  if (tm.val == "inside") {
    tm.delta = -1 * (dt.val / 2);
  } else if (tm.val == "outside") {
    tm.delta = dt.val / 2;
  }

  // drop-down list 2 - horaire (col A) or anti-horaire(col B)
  var cd = {};
  cd.str = "Cutting_Direction";
  cd.val = atelier.getRange("B18").getValue() == "horaire" ? "A" : "B";

  // calculate offset for G41 and G42
  var ofs = {};
  ofs.str = "Offset";
  ofs.val = "()";
  var nofs = {};
  nofs.str = "!Offset"
  nofs.val = "G40";

  // horaire column "A"
  if (cd.val == "A") {
    if (tm.val == "inside"){
      ofs.val = "G42";
    } else if (tm.val == "outside") {
      ofs.val = "G41";
    } else {
      ofs.val = nofs.val = "()";
    }
    // anti-horaire "B" (no other choice)
  } else {
    if (tm.val == "inside"){
      ofs.val = "G41";
    } else if (tm.val == "outside") {
      ofs.val = "G42";
    } else {
      ofs.val = nofs.val = "()";
    }
  }

  // for substitution loop
  var subsA = [
    { str: dt.str, val: dt.val },
    { str: ofs.str, val: ofs.val },
    { str: nofs.str, val: nofs.val },
    { str: "Nb_Pass", val: atelier.getRange("B5").getValue() },
    { str: "D_Cut", val: atelier.getRange("B6").getValue() },
    { str: "Clearance", val: atelier.getRange("B8").getValue() },
    { str: "Spindle_Speed", val: atelier.getRange("B10").getValue() },
    { str: "Rapid_Speed", val: atelier.getRange("B12").getValue() },
    { str: "Low_Speed", val: atelier.getRange("B14").getValue() },
    { str: "Plunge_Speed", val: atelier.getRange("B15").getValue() }
  ];

  // email address to send G codde
  var mail = {};
  mail.str = "Email";
  mail.val = atelier.getRange("B20").getValue();

  // 2. Create Commun strings (sheet2) - overall start and end code 
  // process Gcode arrays to strings with line breaks

  // Debut / start of overall
  var debut = "";
  // get sheet 2 (Commun) column A
  var gCodeA = commun.getRange("A2:A").getValues().filter(String);
  var gLen = gCodeA.length;
  // convert array to string
  for (var k = 0; k < gLen; k++) {
    // first call of substitution function
    // accepts the String to be substituted and the array of vars to check
    var rep = subMe(gCodeA[k].toString(), subsA) + n;
    debut += rep;
  }

  // Fin / very end
  var fin = n;
  gCodeA = commun.getRange("B2:B").getValues().filter(String);
  gLen = gCodeA.length;
  // convert to string
  for (k = 0; k < gLen; k++) {
    rep = subMe(gCodeA[k].toString(), subsA) + n;
    fin += rep;
  }

  // 3. create coupeA array of all valid (non-zero) coupes to select Gcode for sp
  var coupeA = [];
  // currently always 6 values: Type, X, Y, D, Lg, Ht
  var cutLen = 6;
  // get # de coupes
  var b1 = atelier.getRange("B1").getValue();
  // i is rowNum, not an array index!
  for (i = 2; i <= (b1 * 2); i += 2) {
    var temp = "E" + i + ":J" + i;
    var oneCoupe = atelier.getRange(temp).getValues();
    // *** only push if non-zero, skip [0] which is coupe name
    var inclus = false;
    for (var j = 1; j < cutLen; j++) {
      if (oneCoupe[0][j] > 0) {
        inclus = true;
        continue;
      }
    }
    // valid coupe not all zeros, but check for tool diam > circle diam, longueur, larger
    // diameter - dt.val
    var curA = oneCoupe[0];
    if (curA[0]=="Circle" && curA[3] < dt.val){
      SpreadsheetApp.getUi().alert("Le diamètre de la coupe "+i/2+" Circle est plus petit que l'outil !");
      return;
    }
    // rectangle et Lg et Ht
    if ((curA[0]=="Rectangle" && curA[4] < dt.val) || (curA[0]=="Rectangle" && curA[5] < dt.val)){
      SpreadsheetApp.getUi().alert("Le longueur ou hauteur de la coupe "+i/2+" Rectangle est plus petit que l'outil !");
      return;
    }
    if (inclus) {
      // cutsA to store labeled values for each coupe
      // var needs to be refreshed each loop
      var cutVarsA = [
        { str: 'Type', val: "" },
        { str: 'X', val: 0 },
        { str: 'Y', val: 0 },
        { str: 'Dia', val: 0 },
        { str: 'Lg', val: 0 },
        { str: 'Ht', val: 0 }
      ];
      // convert oneCoupe[0] into cutsA, labeling its values
      for (k = 0; k < cutLen; k++) {
        // for each coupe value entered by user
        var cval = oneCoupe[0][k];
        // match with predefined var objs above
        var cut = cutVarsA[k];
        // update Lg, Ht for tool movement using tm.delta before storing
        if ((cut.str == "Lg" || cut.str == "Ht") && (cval>0)) {
          cutDec = cval + tm.delta;
          cut.val = cutDec.toFixed(2);
          continue;
        }
        // store rest
        cut.val = cval;
      }
      // une coupe, X and Y maj, str/val pairs for substition, ready to use
      coupeA.push(cutVarsA);
    }
  }

  // 4. Iterate valid coupes to create middle (milieu)
  var milieu = "";
  // length is num of valid cuts specified
  var len = coupeA.length;
  // for each valid coupe
  for (i = 0; i < len; i++) {
    // show cut number
    milieu += n + "(Coupe " + (i + 1) + " de " + len + ")" + n;
    // spacer between cuts
    milieu += "(##################################)" + n;
    // get the corresponding sheet
    var sheetName = coupeA[i][0].val;
    // get array of values from col A or B based on Cutting Direction (cd.val)
    gCodeA = ss.getRange(sheetName + "!" + cd.val + "2:" + cd.val).getValues().filter(String);
    gLen = gCodeA.length;
    //convert to string
    for (k = 0; k < gLen; k++) {
      // 1. subsA - same for all
      rep = subMe(gCodeA[k].toString(), subsA);
      // 2. coupeA[i] (cutsA) - unique per cut
      rep = subMe(rep, coupeA[i]) + n;
      milieu += rep;
    }
  }

  // program principal
  var pp = debut + milieu + fin;
  // display
  SpreadsheetApp.getUi().alert(pp);
  // sauvegarde Gcode sur Drive
  DriveApp.createFile("LeDernierGcode.nc", pp, MimeType.PLAIN_TEXT);
  //
  // appelle-le
  var file = DriveApp.getFilesByName("LeDernierGcode.nc");
  // send message
  if (file.hasNext()) {
    MailApp.sendEmail(mail.val, 'Votre Gcode - Fraiseuse', pp, {
      attachments: [file.next().getAs(MimeType.PLAIN_TEXT)],
      name: 'automated emailer script'
    });
  }
}

// substitution function - used 4 times by getGcode()
function subMe(lineTxt, subA) {
  var subLen = subA.length;
  for (var i = 0; i < subLen; i++) {
    var tStr = "#" + subA[i].str;
    var tVal = subA[i].val.toString();
    var rep = lineTxt.replace(tStr, tVal);
    lineTxt = rep;
  }
  return lineTxt;
}

/**
 * - creates an open row for each new cut
 * - reads each sheet specified as a type of cut
 * - drop-down menu fit to each row
 * - function does not delete 
 */
function majLatelier() {

  // get all current sheets to get coupes
  var sheetsA = ss.getSheets();

  // get atelier work area
  var atelier = sheetsA[0];

  // update the drop-downs based on all sheets, requires array
  var typeA = [];
  // note - skipping sheetsA[0, 1] l'atelier et commun
  var len = sheetsA.length;
  for (var i = 2; i < len; i++) {
    // make obj and label pairs
    typeA.push(sheetsA[i].getName());
  }
  // set drop-down rule using array of coupe types
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(typeA);
  // get # of coupes
  var b1 = atelier.getRange("B1").getValue();

  // m-a-j all drop-downs 
  var dig = 0;
  for (var i = 0; i < b1; i++) {
    // col E, on even rows 2+
    dig = ((i * 2) + 2);
    atelier.getRange('E'+dig).setDataValidation(rule);

    // col D, add order
    atelier.getRange('D'+dig).setValue(i + 1);

    // colorize rows
    atelier.getRange('D'+dig+':J'+dig).setBackground('#cccccc');
  }
}
