function fetchReward(url, regex) {
  try {
    var options = {
      'muteHttpExceptions': true,
      'headers': { 'Cache-Control': 'no-cache' }
    };
   
    var response = UrlFetchApp.fetch(url, options);
    var html = response.getContentText();
    Logger.log("HTML Content: " + html);
   
    var match = regex.exec(html);
   
    if (match && match[1]) {
      Logger.log("Match Found: " + match[1]);
     
      var value = parseInt(match[1].trim(), 10);
     
      Logger.log("Parsed Reward Value: " + value);
     
      return !isNaN(value) ? value : 0;
    } else {
      Logger.log("No Match Found");
      return 0;
    }
  } catch (e) {
    Logger.log("Error [fetchReward]: " + e.message);
    return 0;
  }
}


function getPoints(url) {
  var regex = /<div class="ygg-overflow-ellipsis">([^<]*)<\/div>/;
  return fetchReward(url, regex);
}


function getYGGToken(url) {
  var regex = /<div class="ygg-overflow-ellipsis">([^<]*)<\/div>/;
  return fetchReward(url, regex);
}


function getPointsPremium(url) {
  var regex = /<div class="YGGPill_icon__xWQRG[^"]*">[\s\S]*?<div class="ygg-overflow-ellipsis">([^<]*)<\/div>/;
  return fetchReward(url, regex);
}

function getYGGTokenPremium(url) {
  var regex = /<div class="YGGPill_icon__xWQRG[^"]*">[\s\S]*?<div class="ygg-overflow-ellipsis">([^<]*)<\/div>/;
  return fetchReward(url, regex);
}


function getQuestManager(url) {
  try {
    var options = {
      'muteHttpExceptions': true,
      'headers': { 'Cache-Control': 'no-cache' }
    };
   
    var response = UrlFetchApp.fetch(url, options);
    var html = response.getContentText();
   
    var regex = /<div class="[^"]*YGGPill_large__RIRGh[^"]*YGGPill_yggPill__gaOdc[^"]*">[\s\S]*?<div class="[^"]*YGGPill_inner__X7z2r[^"]*">[\s\S]*?<div class="[^"]*ygg-overflow-ellipsis[^"]*">([^<]*)<\/div>/;

    var match = regex.exec(html);
   
    if (match && match[1]) {
      return match[1].trim();
    } else {
      return "Unknown";
    }
  } catch (e) {
    return "Unknown";
  }
}


function getQuestCriteria(url){
  try {
    var options = {
      'muteHttpExceptions': true,
      'headers': { 'Cache-Control': 'no-cache' }
    };
   
    var response = UrlFetchApp.fetch(url, options);
    var html = response.getContentText();
   
    var regex = /<div class="[^"]*YGGPill_large__RIRGh[^"]*YGGPill_yggPill__gaOdc[^"]*">[\s\S]*?<div class="[^"]*YGGPill_inner__X7z2r[^"]*">[\s\S]*?<div class="[^"]*ygg-overflow-ellipsis[^"]*">([^<]*)<\/div>/;


    Logger.log(html);
   
    var match = regex.exec(html);
   
    if (match && match[1]) {
     Logger.log("Match: " + match[1]);
      return match[1].trim();
    } else {
     
     Logger.log("Match as: " + match[1]);
      return "Unknown Quest Manager";
    }
  } catch (e) {
    Logger.log("Error: " + e.message);
    return "Unknown Quest Manager";
  }
}


function getQuestName(url) {
  if(url == ""){
      return "No URL Provided";
    } else{
      var options = {
      'muteHttpExceptions': true,
      'headers': { 'Cache-Control': 'no-cache' } // Disable caching
    };
   
    var response = UrlFetchApp.fetch(url, options);
      var html = response.getContentText();
     
      // Use regex to find the content of the <title> tag
      var titleMatch = html.match(/<title>(.*?)<\/title>/);


      if (titleMatch && titleMatch[1]) {
          return titleMatch[1]; // title text
      } else {
          return "No title tag found";
      }
    }
}


function questTypeSelected(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
 
  var column = range.getColumn();
  var editedColumnA = 1;
  var dropdownColumnB = 2;
  var dropdownColumnC = 3;


  if (column === editedColumnA || column === dropdownColumnB || column === dropdownColumnC) {
    var url = sheet.getRange(row, editedColumnA).getValue();
    var selectedValue = sheet.getRange(row, dropdownColumnB).getValue();
    var selectedPremium = sheet.getRange(row, dropdownColumnC).getValue();


    sheet.getRange(row, 5).clearContent();
    sheet.getRange(row, 6).clearContent();


    if(selectedPremium === "Premium" && selectedValue === "Game") {
        sheet.getRange(row, 5).setValue(getPointsPremium(url));
        sheet.getRange(row, 6).setValue(0);
    } else if (selectedPremium === "Premium" && selectedValue === "Bounty") {
        sheet.getRange(row, 5).setValue(0);
        sheet.getRange(row, 6).setValue(getYGGTokenPremium(url));
    } else if (selectedPremium === "Non-Premium" && selectedValue === "Game") {
        sheet.getRange(row, 5).setValue(getPoints(url));
        sheet.getRange(row, 6).setValue(0);
    } else if (selectedPremium === "Non-Premium" && selectedValue === "Bounty") {
        sheet.getRange(row, 5).setValue(0);
        sheet.getRange(row, 6).setValue(getYGGToken(url));
    }
  }
}


function isQuestPremium(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
}


function clearFields(e){
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();


  var column = range.getColumn();
  var editedColumnA = 1;
  var dropdownColumnB = 2;
  var dropdownColumnC = 3;


   if (column === editedColumnA || column === dropdownColumnB) {
    var url = sheet.getRange(row, editedColumnA).getValue();
    var selectedValue = sheet.getRange(row, dropdownColumnB).getValue();
    var selectedPremium = sheet.getRange(row, dropdownColumnC).getValue();
   
    if(url === ""){
      sheet.getRange(row, 5).clearContent();
      sheet.getRange(row, 6).clearContent();
      sheet.getRange(row, 2).clearContent();
      sheet.getRange(row, 3).clearContent();
    }
   }
}


function checkEmptyCellsInColumn(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnA = 1;
  var lastRow = sheet.getLastRow();
 
  for (var row = 1; row <= lastRow; row++) {
    var cellValueA = sheet.getRange(row, columnA).getValue();
   
    if (!cellValueA) { 
      sheet.getRange(row, 5).clearContent();
      sheet.getRange(row, 6).clearContent();
      sheet.getRange(row, 2).clearContent();
      sheet.getRange(row, 3).clearContent();
    }
  }
}


function onEdit(e) {
  var sheet = e.source.getActiveSheet();
 
  var targetSheetName = "Quests";

  if (sheet.getName() !== targetSheetName) {
    return;
  }


  questTypeSelected(e);
  clearFields(e);
  checkEmptyCellsInColumn(e);
}







