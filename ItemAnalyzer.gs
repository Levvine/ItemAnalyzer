function ItemAnalyzer() {
  
  var activeSheetName = "ItemCalc";
  
  var spreadsheet = SpreadsheetApp.getActive();
  var interface = spreadsheet.getSheetByName("Interface");
  var items = spreadsheet.getSheetByName("Items");
  var data = spreadsheet.getSheetByName(activeSheetName);
  var presentation = spreadsheet.getSheetByName("Presentation");
  
  Logger.log("Start Debug Log:");
  
  // clears data sheet, resets interface
  if (data != null) {
    data.clear();
  } else {
    data = spreadsheet.insertSheet(activeSheetName);
  }
  
  reset();
  
  
  /* INIT */
  
  
  // gets ranges from interface
  // enemy stats
  var eHP = interface.getRange("O25");
  var eAR = interface.getRange("O26");
  var eMR = interface.getRange("O27");
  
  /*
  // champion, level
  var champ = interface.getRange("B3");
  var lv = interface.getRange("H3");
  */
  
  // skillpoints
  var rankQ = interface.getRange("L4");
  var rankW = interface.getRange("L5");
  var rankE = interface.getRange("L6");
  var rankR = interface.getRange("L7");
  
  // backs up enemy stats
  var eHP_bak = eHP.getFormula();
  var eAR_bak = eAR.getFormula();
  var eMR_bak = eMR.getFormula();
  
  // gets range of preset enemy stats
  var eStat = presentation.getRange(4, 16, 6, 9); // P4:X9
  
  // note: runtime error occurs when a mythic item is in a non-primary slot
  // note: runtime error occurs when there is more than one mythic item
  // builds
  var firstB = [
    ["Long Sword", "Noonquiver"],
    ["Long Sword", "Long Sword", "Long Sword", "Cloak of Agility"],
    ["Long Sword", "Long Sword", "Long Sword", "Dagger", "Boots of Speed"],
    ["Berserker's Greaves", "Long Sword", "Long Sword"],
    ["Caulfield's Warhammer", "Tear of the Goddess", "Boots of Speed"],
    ["Pickaxe", "Vampiric Scepter"]
  ];
  
  var secondB = [
    ["Immortal Shieldbow", "Boots of Speed"],
    ["Galeforce", "Boots of Speed"],
    ["Kraken Slayer", "Boots of Speed"],
    ["Manamune", "Ionian Boots of Lucidity", "Faerie Charm"],
    ["Blade of the Ruined King", "Dagger", "Boots of Speed"]
  ];
  
  var thirdB = [
    ["Kraken Slayer", "Berserker's Greaves", "Cloak of Agility", "Dagger", "Dagger"],
    ["Kraken Slayer", "Berserker's Greaves", "Rageknife", "Dagger"],
    ["Kraken Slayer", "Berserker's Greaves", "Pickaxe", "Long Sword"],
    ["Manamune", "Ionian Boots of Lucidity", "Bandleglass Mirror", "Kindlegem", "Long Sword"],
    ["Blade of the Ruined King", "Berserker's Greaves", "Noonquiver"]
    ];
    
  // builds is a 3D array of build costs - builds - items
  var builds = [firstB, secondB, thirdB];
  var maxCost = [1800, 3800, 5800, 7800, 9800, 11800];
  var lv = [6, 9, 11, 13, 15, 17];
  var skillpoint = ["W", "E", "W", "Q", "W", "R", "W", "Q", "W", "Q", "R", "Q", "Q", "E", "E", "R", "E", "E"];
  var champ = "Ashe";
  
  
  /* OUTPUT */
  
  /* 
   * Spacing Notes
   * Build Table: 12 spaces
   */
  
  // sets table offsets
  // var rowsOffsetTable = 0;
  // var colsOffsetTable = 0;
    
  prevLv = 0; // initial level 0 for skillpoint increment
  
  // sets champion name
  interface.getRange("B3").setValue(champ);
  
  // sets damage type to rotation
  interface.getRange("H13").setValue("Trade");
    
  // create tables
  // iterate through build costs 
  for (var i = 0; i < builds.length; i++) {
    
    // interface - enter level and skillpoints
    interface.getRange("H3").setValue(lv[i]);
    for (var m = prevLv; m < lv[i]; m++) {
      switch (skillpoint[m]) {
        case "Q":
          rankQ.setValue(rankQ.getValue() + 1);
          break;
        case "W":
          rankW.setValue(rankW.getValue() + 1);
          break;
        case "E":
          rankE.setValue(rankE.getValue() + 1);
          break;
        case "R":
          rankR.setValue(rankR.getValue() + 1);
      }
    }
    
    // create table headings
    data.getRange(1, i * 13 + 1).setValue(maxCost[i] + " gold (lv"+lv[i]+")");
    data.getRange(2, i * 13 + 7).setValue("All-in (sec.)");
    data.getRange(2, i * 13 + 10).setValue("Rotation (%HP)");
    data.getRange(2, i * 13 + 1).setValue("Build");
    
    // create table subheadings
    data.getRange(3, i * 13 + 7).setValue("Tank");
    data.getRange(3, i * 13 + 8).setValue("Bruiser");
    data.getRange(3, i * 13 + 9).setValue("Squish");
    data.getRange(3, i * 13 + 10).setValue("Tank");
    data.getRange(3, i * 13 + 11).setValue("Bruiser");
    data.getRange(3, i * 13 + 12).setValue("Squish");
    
    // note: logic error occurs if more than one instance of item
    // iterate through builds
    var textFinder;
    var icon;
    for (var j = 0; j < builds[i].length; j++) {
      
      // iterate through items
      for (var k = 0; k < builds[i][j].length; k++) {
        
        // writes item icons to data
        textFinder = items.createTextFinder(builds[i][j][k]);
        var match = textFinder.findNext();
        icon = items.getRange(match.getRow(), 40).getFormula();
        icon = icon.slice(0,-1) + ",4,36,36)";
        data.getRange(j * 3 + 4, i * 13 + k + 1).setFormula(icon);
        
        // changes item slots in interface
        var item = builds[i][j][k];
        interface.getRange(22 + k, 3).setValue(item);
        
      } // end item loop
      
      // iterate through enemy archetypes, 3 columns: HP, AR, MR
      for (var l = 0; l < eStat.getNumColumns() / 3; l++) {
        
        // seed enemy stats by build cost and archetype
        eHP.setValue(eStat.getCell(i + 1, l * 3 + 1).getValue());
        eAR.setValue(eStat.getCell(i + 1, l * 3 + 2).getValue());
        eMR.setValue(eStat.getCell(i + 1, l * 3 + 3).getValue());
        
        // writes ttk (all in) to data
        data.getRange(j * 3 + 4, i * 13 + l + 7).setValue(interface.getRange("F13").getValue()).setNumberFormat("0.0");
        
        // todo: write trade (rotation) to data
        data.getRange(j * 3 + 4, i * 13 + l + 10).setValue(interface.getRange("I13").getValue() / eHP.getValue()).setNumberFormat("0.0%");
        
      } // end enemy archetype loop
      
      resetItems();
      prevLv = lv[i];
      
    } // end build loops
    
  } // end build cost loops

  
  /* CLEANUP */
  reset();
  
}

function createEnemies() {
  
  var activeSheetName = "EnemyCalc";
  
  var spreadsheet = SpreadsheetApp.getActive();
  var interface = spreadsheet.getSheetByName("Interface");
  var items = spreadsheet.getSheetByName("Items");
  var data = spreadsheet.getSheetByName(activeSheetName);
  var presentation = spreadsheet.getSheetByName("Presentation");
  
  Logger.log("Start Debug Log:");
  
  // clears data sheet, resets Interface
  if (data != null) {
    data.clear();
  } else {
    data = spreadsheet.insertSheet(activeSheetName);
  }
  
  reset();
  
  
  /* INIT */
  
  var firstBTank = ["Bami's Cinder", "Boots of Speed", "Cloth Armor"];
  var firstBBruiser = ["Ironspike Whip", "Boots of Speed", "Cloth Armor"];
  var firstBSquishy = ["Tear of the Goddess", "Sheen", "Long Sword", "Long Sword"];
  
  var secondBTank = ["Sunfire Aegis", "Boots of Speed", "Cloth Armor"];
  var secondBBruiser = ["Ironspike Whip", "Phage", "Kindlegem", "Boots of Speed", "Cloth Armor"];
  var secondBSquishy = ["Manamune", "Sheen", "Boots of Speed"];
  
  var thirdBTank = ["Sunfire Aegis", "Plated Steelcaps", "Bramble Vest", "Ruby Crystal"];
  var thirdBBruiser = ["Goredrinker", "Plated Steelcaps", "Caulfield's Warhammer"];
  var thirdBSquishy = ["Manamune", "Ionian Boots of Lucidity", "Sheen", "Phage", "Ruby Crystal"];
  
  var buildsTank = [firstBTank, secondBTank, thirdBTank];
  var buildsBruiser = [firstBBruiser, secondBBruiser, thirdBBruiser];
  var buildsSquishy = [firstBSquishy, secondBSquishy, thirdBSquishy];
  var builds = [buildsTank, buildsBruiser, buildsSquishy];
  
  var champTank = "Sion";
  var champBruiser = "Renekton";
  var champSquishy = "Ezreal";
  var champs = [[champTank, "Tank"], [champBruiser, "Bruiser"], [champSquishy, "Squishy"]];
  var stats = ["HP", "AR", "MR"];

  // champs = ["name", "archetype", "exp type", [builds]];  look into javascript Map
  // js map good for unique keys but slower to find index than arrays? at least, slower to write.
  
  var maxCost = [1800, 3800, 5800, 7800, 9800, 11800];   
     
  var lvSolo = [6, 9, 11, 13, 15, 17];
  var lvDuo;
  

  /* OUTPUT */
  
  data.getRange("A2").setValue("Gold");
  
  // i = index of build costs, or recall
  for (var i = 0; i < builds.length; i++) {
    
    Logger.log("i = " + i);
    Logger.log("maxCost = " + maxCost[i]);
    
    data.getRange(i + 3, 1).setValue(maxCost[i]);
    
    // j = index of champ archetypes
    for (var j = 0; j < builds.length; j++) {
      
      // headings
      // champs [i][0] = name, champs[i][1] = archetype
      data.getRange(1, j * stats.length + 2).setValue(champs[j][1]);
      data.getRange(2, j * stats.length + 2, 1, stats.length).setValues([stats]);
      
      // sets champion name and level in interface
      interface.getRange("N16").setValue(champs[j][0]);
      interface.getRange("O17").setValue(lvSolo[i]);
            
      // k = index of items
      for (var k = 0; k < builds[j][i].length; k++) {        
        interface.getRange(k + 18, 15).setValue(builds[j][i][k]);
      }
      
      // copy stats from interface to data
      data.getRange(i + 3, j * stats.length + 2).setValue(interface.getRange("O29").getValue()).setNumberFormat("0");
      data.getRange(i + 3, j * stats.length + 3).setValue(interface.getRange("O30").getValue()).setNumberFormat("0");
      data.getRange(i + 3, j * stats.length + 4).setValue(interface.getRange("O31").getValue()).setNumberFormat("0");
      
      // reset Enemy field
      resetEnemy();
    }
      
  }
  
  
  /* CLEANUP */
  reset();
  
}
