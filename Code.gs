/**
*
* Primary function, get things started. The purpose of this script is to return your Dropbox links from https://www.dropbox.com/share/links to a Google Sheet.
*
*/

function primaryFunction(){
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get shared Dropbox files
  getSharedDropboxFiles(spreadsheet);
}

/**
*
* Link to Dropbox API and return shared files.
*
* Directions
* 1. Go to https://www.dropbox.com/developers/apps/create?_tk=pilot_lp&_ad=ctabtn1&_camp=create
* 2. Select Dropbox API
* 3. Select Full Dropbox
* 4. Give your App a name (I gave it ryanmcslomo-GoogleAppsScript)
* 5. On the next screen, generate an access token and save it to var dropboxAccessToken on line 30
*
* References
*
* https://www.labnol.org/code/20571-download-web-files-dropbox
* https://www.labnol.org/code/20573-upload-google-drive-files-dropbox
* https://www.dropbox.com/developers/documentation/http/documentation#sharing-list_shared_links
*
* @param spreadsheet {Object} The active spreadsheet object. This is where we'll print the array.
* @param cursor {String} The cursor returned by your last call to list_shared_links, indicates our position in returning links.
*
*/

function getSharedDropboxFiles(spreadsheet, cursor) {
  
  //  Pause script to not trigger API limits
  Utilities.sleep(3000);
  
  //  Declare variables
  var linkArray = [];
  var parameters = {
    // This is optional. You can delete this and return all shared files or add paths to items or you can add paths to folders. For example:    
    //    "path": "/graduate school/ryan's uploads"
    //    "path": "/graduate school/ryan's uploads/picture.jpg"
    // The slashes indicate folder hierarchy. You can also use path ID and a few other tricks.
    // More info: https://www.dropbox.com/developers/documentation/http/documentation#sharing-list_shared_links
  };
  
  if (cursor){
    parameters.cursor = cursor;
  }
  
  // Add your Dropbox Access Token
  var dropboxAccessToken = 'xXxXxXxXx';
  
  //  Set authentication object parameters
  var headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + dropboxAccessToken,
  };
  
  //  Set option parameters
  var options = {
    "method": "POST",
    "headers": headers,
    "muteHttpExceptions": true,
//    "payload": JSON.stringify(parameters)
  };
  
  //  Hit up API
  var apiUrl = "https://api.dropboxapi.com/2/sharing/list_shared_links";
  try{
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseText = response.getContentText();
    var responseTextJSON = JSON.parse(responseText);
    
    //  Parse JSON response
    var links = responseTextJSON.links;
    var hasMore = responseTextJSON.has_more;
    var cursor = responseTextJSON.cursor;  
    for (var link = 0; link < links.length; link++){
      linkArray.push([links[link].name, links[link].path_lower, links[link].id, links[link][".tag"], links[link].url]);    
    }
  } catch (e) {
    console.log(e);
    linkArray.push(e);
  }
  
  //  Print to sheet and continue if there are still more entries  
  setArrayValuesToSheet(spreadsheet, linkArray, hasMore, cursor);  
} 

/**
*
* Print array to sheet.
*
* @param sheet {Object} The active spreadsheet object. This is where we'll print the array.
* @param linkArray {Array} The array of returned Dropbox items.
* @param hasMore {Boolean} True if there are more entries, false if we have grabbed them all.
* @param cursor {String} The cursor returned by your last call to list_shared_links, indicates our position in returning links.
*
*/

function setArrayValuesToSheet(spreadsheet, linkArray, hasMore, cursor){
  
  //  Add header row if not present
  var spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var firstCell = sheet.getRange(1, 1).getValue();
  var lastRow = sheet.getLastRow();
  if (firstCell != 'Name' && linkArray.length > 1) {
    var headerRow = ["Name", "Path", "ID", "Tag", "URL"];
    linkArray.unshift(headerRow);
  }
  
  //  Print error message if we got one
  if (linkArray.length === 1){
    sheet.getRange(lastRow + 1, 1).setValue(linkArray); 
  }
  
  //  Print array to active sheet  
  sheet.getRange(lastRow + 1, 1, linkArray.length, linkArray[0].length).setValues(linkArray); 
  SpreadsheetApp.flush();
  
  //  If there are more Dropbox files, run function again
  if (hasMore){
    getSharedDropboxFiles(spreadsheet, cursor);
  }
}


/**
*
* Create a menu option for script functions.
*
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
  .addItem('Get Shared Dropbox Files', 'primaryFunction')
  .addToUi();  
}
