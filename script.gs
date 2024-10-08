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
      'headers': { 'Cache-Control': 'no-cache' }
    };
    
    var response = UrlFetchApp.fetch(url, options);
      var html = response.getContentText();
      
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
      Logger.log("Cleared column D for row: " + row); 
    }
  }
}

function magdagsdiscordClicked() {
  var html = "<script>window.open('https://discord.com/users/478517550431338507');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html); 
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Loading...');
}

function magdagstwitterClicked() {
  var html = "<script>window.open('https://twitter.com/mangdags');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html); 
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Loading...');
}

function yggLogoClicked() {
  var html = "<script>window.open('https://discord.gg/ygg');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html); 
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Loading...');
}


function roninClicked(){
  var sheet = SpreadsheetApp.getActiveSheet();

  var qrBase64 = "data:image/png;base64,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";

  var html = '<div style="display: flex; justify-content: center; align-items: center; height: 100%;">' +
               '<img src="' + qrBase64 + '" style="max-width: 100%; max-height: 100%;">' +
             '</div>';
  var ui = SpreadsheetApp.getUi();

  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(300), 'MangDags Ronin Wallet');

}

function forceRefresh() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var targetSheetName = "Tracker";

  if (sheet.getName() !== targetSheetName) {
    return;
  }

  questTypeSelected({ source: SpreadsheetApp.getActiveSpreadsheet(), range: sheet.getActiveRange() });
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  
  var targetSheetName = "Tracker";

  if (sheet.getName() !== targetSheetName) {
    return;
  }

  questTypeSelected(e);
  clearFields(e);
  checkEmptyCellsInColumn(e);
}


