function fetchReward(url, regex) {
  try {
    var options = {
      'muteHttpExceptions': true,
      'headers': { 'Cache-Control': 'no-cache' } // Disable caching
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var html = response.getContentText();
    Logger.log("HTML Content: " + html);
    
    // Execute the regex to find the first match
    var match = regex.exec(html);
    
    // Log the match for debugging
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
      'headers': { 'Cache-Control': 'no-cache' } // Disable caching
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var html = response.getContentText();
    
    var regex = /<div class="[^"]*YGGPill_large__RIRGh[^"]*YGGPill_yggPill__gaOdc[^"]*">[\s\S]*?<div class="[^"]*YGGPill_inner__X7z2r[^"]*">[\s\S]*?<div class="[^"]*ygg-overflow-ellipsis[^"]*">([^<]*)<\/div>/;


    // Log the raw HTML for debugging
    Logger.log(html);
    
    var match = regex.exec(html);
    
    if (match && match[1]) {
     Logger.log("Match: " + match[1]);
      return match[1].trim(); 
    } else {
      
     Logger.log("Match as: " + match[1]);
      return "Unknown";
    }
  } catch (e) {
    Logger.log("Error: " + e.message);
    return "Unknown"; // Return default value in case of an error
  }
}

function getQuestCriteria(url){
  try {
    var options = {
      'muteHttpExceptions': true,
      'headers': { 'Cache-Control': 'no-cache' } // Disable caching
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var html = response.getContentText();
    
    var regex = /<div class="[^"]*YGGPill_large__RIRGh[^"]*YGGPill_yggPill__gaOdc[^"]*">[\s\S]*?<div class="[^"]*YGGPill_inner__X7z2r[^"]*">[\s\S]*?<div class="[^"]*ygg-overflow-ellipsis[^"]*">([^<]*)<\/div>/;

    Logger.log(html);
    
    var match = regex.exec(html);
    
    if (match && match[1]) {
     Logger.log("Match: " + match[1]);
      return match[1].trim(); // Return the quest manager's name
    } else {
      
     Logger.log("Match as: " + match[1]);
      return "Unknown Quest Manager";
    }
  } catch (e) {
    Logger.log("Error: " + e.message);
    return "Unknown Quest Manager"; // Return default value in case of an error
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Get active sheet
  var columnA = 1; // Column A (1 represents column A)
  var lastRow = sheet.getLastRow(); // Get the last row with content in the sheet
  
  // Loop through each row in column A
  for (var row = 1; row <= lastRow; row++) {
    var cellValueA = sheet.getRange(row, columnA).getValue(); // Get the value of the cell in column A
    
    // If column A is empty, clear the corresponding cell in column D
    if (!cellValueA) { // This condition checks if the cell in column A is empty
      sheet.getRange(row, 5).clearContent(); 
      sheet.getRange(row, 6).clearContent(); 
      sheet.getRange(row, 2).clearContent();
      sheet.getRange(row, 3).clearContent();
      Logger.log("Cleared column D for row: " + row); // Log the action for debugging
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


function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  
  // Specify the name of the sheet you want to target
  var targetSheetName = "Quests";

  // Check if the active sheet is the one you want to work with
  if (sheet.getName() !== targetSheetName) {
    return; // Exit the function if the sheet is not the target sheet
  }

  questTypeSelected(e);
  clearFields(e);
  checkEmptyCellsInColumn(e);
}


