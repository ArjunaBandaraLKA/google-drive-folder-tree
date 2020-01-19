/*
 * 
 * Google Apps Script - List all files & folders in a Google Drive folder, & write into a speadsheet.
 * Hint: Set your folder ID first! You may copy the folder ID from the browser's address field. 
 *       The folder ID is everything after the 'folders/' portion of the URL.
 * 
 * @Open source - Original idea from https://github.com/mesgarpour.
 * @Modified by Wenjun Zhang: zw.zhangwenjun@gmail.com 電話：九二五，三二九，八七四五
 * 
 * @version 3 (updated 8/7/2018)
 * Add function to continue with the last left over. 
 *
 * 
 */
// ======================================================================================================
function onOpen() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('List Files & Attributes')
    .addItem('List All Files', 'listAll')
    .addSeparator()
    .addSubMenu(ui.createMenu('List Attributes (Optional)')
      .addItem('Owner', 'ListOwner')
      .addItem('Editors', 'ListEditors')
      .addItem('Viewers', 'ListViewers')
      .addItem('Description', 'ListDescription')
      .addItem('Date Created', 'ListDateCreated')
      .addItem('Last Updated', 'ListLastUpdated')
      .addItem('Size', 'ListSize')
      .addItem('File Sharing status', 'ListAccess')
      .addItem('User permission', 'ListPermission')
    )
    .addToUi();
}

function listAll() {
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var lastRow = currentSheet.getLastRow();
  try{
    var temp = currentSheet.getRange(lastRow,1).getValue();}
  catch (e) {
    var temp = '0';}
  
  var marker = 'All files has been listed by ';
  if (temp.indexOf(marker) > -1 ) {
    var userInput = Browser.msgBox('Process - Question', 'Seems like ' + temp + '. Click "YES" to start a new list, Click "NO" to cancel and keep the current.', Browser.Buttons.YES_NO);
    if (userInput == "no") {return;}
    
  }
  else {
    var userInput = Browser.msgBox('Process - Question', 'Start New List OR Continue with existing? Click "YES" to start new, Click "NO" to continue with existing sheet.', Browser.Buttons.YES_NO);}
  if (userInput == "yes") {
    currentSheet.clear();
    var folderId = Browser.inputBox('Enter folder ID', Browser.Buttons.OK_CANCEL);
    currentSheet.appendRow(["Gdrive ID", "Name", "Full Path", "URL", folderId]);
//    currentSheet.getRange(2, 6).setFormula('ARRAYFORMULA(if(E2:E="Folder",A2:A,""))');
    var list = [];
    var excluded = [];
    //var folderId = Browser.inputBox('Enter folder ID or folder link:', Browser.Buttons.OK_CANCEL);
    //folderId = folderId.toString().match(/[-\w]{25,}/);
    //----------Set folder ID, suggest manually input if you have a large list and will run this script several times

    if (folderId === "") {
      Browser.msgBox('Invalid folder ID');
      return;
    }
  } //------- Make sure FolderID is not null. (If folder ID is null, it will excute as the personal root folder) Room to improve here. 
  else if (userInput == "no") {
    //       var voidFolder = Browser.inputBox('Enter sub-folder IDs you would like to skip. (If you have more than one to skip, type "," between IDs, Leave empty if none): ', Browser.Buttons.OK_CANCEL);
    temp = currentSheet.getRange(1, 5, lastRow).getValues();
    var folderId = temp[0][0];
    var list = cleanArray([].concat.apply([], temp));
    var lastID = list.pop();
    var excluded = [lastID,];
    getVoidFolderList(lastID, list, excluded);
  } else {
    return;
  }


  var parent = DriveApp.getFolderById(folderId);
  var parentName = DriveApp.getFolderById(folderId).getName();
  getChildFolders(parentName, parent, currentSheet, list, excluded);
  getRootFiles(parentName, parent, currentSheet);
  SpreadsheetApp.setActiveSheet(currentSheet).getRange(currentSheet.getLastRow()+1,1).setValue(marker + new Date()).setFontColor("#ff0000").setBackground("#ffff00"); //set a marker once completed                 
  SpreadsheetApp.flush();
  Browser.msgBox(marker + new Date());
}
// ======================================================================================================

function cleanArray(actual) { // Will remove all falsy values: undefined, null, 0, false, NaN and "" (empty string)
  var newArray = new Array();
  for (var i = 0; i < actual.length; i++) {
    if (actual[i]) {
      newArray.push(actual[i]);
    }
  }
  return newArray;
}
//=====================================================================================================
/* Not in use.
// ----------Get Folder Tree
function getFolderTree(folderId, voidFolder) {
    try {
        // ----------Get folder by id
        var parentFolder = DriveApp.getFolderById(folderId);
        // ----------Initialise the sheet
//Need to add whether this is new;      
        // ----------Get files and folders
        getChildFolders(parentFolder.getName(), parentFolder, data, sheet, voidFolder);

    } catch (e) {
        Logger.log(e.toString()); //--------------catch error
    }
  
    getRootFiles(parentFolder.getName(), parentFolder, data, sheet);
    Browser.msgBox('List all completed');
};
*/

// ======================================================================================================
// ----------code below is function to get root file names.

function getRootFiles(parentName, parent, sheet) {
  var files = parent.getFiles();
  var output = [];
  var path;
  var Url;
  var fileID;
  var FileOwnerEmail;
  while (files.hasNext()) {
    var childFile = files.next();
    var fileName = childFile.getName();
    path = parentName + " |--> " + fileName;
    //       fileName,
    //       childFile.getDateCreated(),
    //       childFile.getLastUpdated(),
    //       childFile.getDescription(),
    //       childFile.getSize(),
    //     FileOwnerEmail=childFile.getOwner().getEmail();
    //        childFile.getParents().getUrl()
    fileID = childFile.getId();
    Url = "https://drive.google.com/open?id=" + fileID;
    // ---------- Write
    output.push([fileID, fileName, path, Url]);
  }
  if (output.length) {
    var last_row = sheet.getLastRow();
    sheet.getRange(last_row + 1, 1, output.length, 4).setValues(output);
  }
  /*
      while (rootFiles.hasNext()) {
          var rootFile = rootFiles.next();
          var fileId = rootFile.getId();
          var fileName = rootFile.getName();
          var data = [
              fileId,
              fileName,
              parentName + " |--> " + fileName,
              //      rootFile.getDateCreated(), // ----------cut out to save time
              ("https://drive.google.com/open?id=" + fileId)
              //      rootFile.getLastUpdated(), // ----------cut out to save time
              //      rootFile.getDescription(), // ----------cut out to save time
              //      rootFile.getSize(), // ----------cut out to save time
              //      rootFile.getOwner().getEmail()
              //       rootFile.getParents().getUrl() // ----------cut out to save time
          ];
          // ---------- Write
          output.push([fileID, fileName, path, Url]);
          sheet.appendRow(data);
      }
  */
}
// ======================================================================================================

function getChildFolders(parentName, parent, sheet, voidFolder, excluded) {
  var childFolders = parent.getFolders();

  // ---------- List folders inside the folder
  while (childFolders.hasNext()) {
    var childFolder = childFolders.next();
    var folderID = childFolder.getId();

    if (voidFolder.indexOf(folderID) > -1) { // the folder and files in it has been listed
      continue;
    }

    var folderName = childFolder.getName();

    var data = [
      folderID,
      folderName,
      parentName + " |--> " + folderName,
      //     childFolder.getDateCreated(), // ----------cut out to save time
      ("https://drive.google.com/open?id=" + folderID),
      //     childFolder.getLastUpdated(), // ----------cut out to save time
      //     childFolder.getDescription(), // ----------cut out to save time
      //     childFolder.getSize(),  // ----------cut out to save time
      //     childFolder.getOwner().getEmail(),
      folderID, // ----------indicate this is a folder 
    ];
    // ---------- Write
    if (excluded.indexOf(folderID) == -1) { //check the situation of folder is listed but the files have not;
      sheet.appendRow(data);
    }

    // ---------- List files inside the folder
    var files = childFolder.getFiles();
    var output = [];
    var path;
    var Url;
    var fileID;
    var FileOwnerEmail;
    while (files.hasNext()) {
      var childFile = files.next();
      var fileName = childFile.getName();
      path = parentName + " |--> " + folderName + " |--> " + fileName;
      //       fileName,
      //       childFile.getDateCreated(),
      //       childFile.getLastUpdated(),
      //       childFile.getDescription(),
      //       childFile.getSize(),
      //     FileOwnerEmail=childFile.getOwner().getEmail();
      //        childFile.getParents().getUrl()
      fileID = childFile.getId();
      Url = "https://drive.google.com/open?id=" + fileID;
      // ---------- Write
      output.push([fileID, fileName, path, Url]);
    }
    if (output.length) {
      var last_row = sheet.getLastRow();
      sheet.getRange(last_row + 1, 1, output.length, 4).setValues(output);
    }

    // ---------- Recursive call of the subfolder
    getChildFolders(parentName + " |--> " + folderName, childFolder, sheet, voidFolder, excluded);

  }
}

// ======================================================================================================
// ---------- Code below is to exclude the parent ID where the script stops so that it will catch the folder from beginning. 


function getVoidFolderList(lastID, list, excluded) {

  var folderParents = DriveApp.getFolderById(lastID).getParents();
  var folderID = new Array;

  while (folderParents.hasNext()) {
    var folderID = folderParents.next().getId();
    if (folderID == "NOT FOUND") {
      break;
    }

    var index = list.indexOf(folderID);
    if (index !== -1) {
      excluded.push(folderID);
      list.splice(index, 1);
    }

    getVoidFolderList(folderID, list, excluded);
  }
}

/*
{
    var folderChild = Browser.inputBox('Enter the folder ID that have not completed', Browser.Buttons.OK_CANCEL);
    var i = 1;
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = ss.getLastRow();
    var cell = ss.getRange(lastRow + 1, 2)
    cell.setValue(folderChild);
    getParentFolderID(folderChild, i, lastRow);
}

function getParentFolderID(folderChild, i, lastRow) {

    var folderParents = DriveApp.getFolderById(folderChild).getParents();
    var folderID = new Array;

    while (folderParents.hasNext()) {
        var folderID = folderParents.next().getId();
        if (folderID == "NOT FOUND") {
            break;
        }
        i++;
        var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var cell = ss.getRange(lastRow + 1, i + 1);
        cell.setValue(folderID);

        getParentFolderID(folderID, i, lastRow);
    }
}
*/
// ======================================================================================================


function ListOwner() {

    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastrow = ss.getLastRow();
    var lastcolumn = ss.getLastColumn();
    var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

    while (startnumber == 1) {
        startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
    }

    ss.getRange(1, lastcolumn + 1).setValue("Owner");

    for (var x = startnumber; x <= lastrow; x++)

    {
        var value = ss.getRange(x, 1).getValue();
        if (value == "") {
            continue;
        }
        try {
            var childFolder = DriveApp.getFolderById(value);
        } catch (e) {
            ss.getRange(x, lastcolumn + 1).setValue("Fail");
            continue;
        }


        try {
            var owner = childFolder.getOwner().getEmail();
            ss.getRange(x, lastcolumn + 1).setValue(owner);
        } catch (e) {
            ss.getRange(x, lastcolumn + 1).setValue("team drive");
        }
    }
}

//===========================================================================================================


function ListEditors() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue("Editors");

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    var editors = childFolder.getEditors();
    var editorValues = [];

    for (var v = 0; v < editors.length; v++) {
      editorValues.push(editors[v].getEmail());
    }

    var joinValues = editorValues.join(", ");
    ss.getRange(x, lastcolumn + 1).setValue(joinValues);


    if (joinValues == "") {
      ss.getRange(x, lastcolumn + 1).setValue("-");
    }
  }
}

//===========================================================================================================

function ListViewers() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue("Viewers");

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    var viewers = childFolder.getViewers();
    var viewerValues = [];

    for (v = 0; v < viewers.length; v++) {
      viewerValues.push(viewers[v].getEmail());

    }

    var joinValues = viewerValues.join(", ");
    ss.getRange(x, lastcolumn + 1).setValue(joinValues);

    if (joinValues == "") {
      ss.getRange(x, lastcolumn + 1).setValue("-");
    }


  }

}

//===========================================================================================================

function ListDescription() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue("Description");

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    var Description = childFolder.getDescription();
    ss.getRange(x, lastcolumn + 1).setValue(Description);
    if (Description == "" || Description == "undefined" || Description == null) {
      ss.getRange(x, lastcolumn + 1).setValue("-");
    }

  }

}

//==================================================================================================================

function ListDateCreated() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue("Date Created");

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    var DateCreated = childFolder.getDateCreated();
    ss.getRange(x, lastcolumn + 1).setValue(DateCreated);

  }

}

//========================================================================================================================

function ListLastUpdated() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue("Last Updated");

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    var LastUpdated = childFolder.getLastUpdated();
    ss.getRange(x, lastcolumn + 1).setValue(LastUpdated);

  }

}

//===============================================================================================================================

function ListSize() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue("Size");

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    var Size = childFolder.getSize();
    ss.getRange(x, lastcolumn + 1).setValue(Size);

    if (Size == "0") {
      ss.getRange(x, lastcolumn + 1).setValue("-");
    }

  }
}

//================================================================================================================================


function ListAccess() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue('=HYPERLINK("https://developers.google.com/apps-script/reference/drive/access","Who Can Access")');

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    try {
      var sharingAccess = childFolder.getSharingAccess();
      ss.getRange(x, lastcolumn + 1).setValue(sharingAccess);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

  }
}

//===========================================================================================================================================

function ListPermission() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();
  var startnumber = Browser.inputBox('Please input row (number) you would like to be start with. Note: Do NOT put "1" as it is headline', Browser.Buttons.OK_CANCEL);

  while (startnumber == 1) {
    startnumber = Browser.inputBox('Row 1 is headline. Pick another row number', Browser.Buttons.OK_CANCEL);
  }

  ss.getRange(1, lastcolumn + 1).setValue('=HYPERLINK("https://developers.google.com/apps-script/reference/drive/permission","User Permission to file/folder")');

  for (var x = startnumber; x <= lastrow; x++)

  {
    var value = ss.getRange(x, 1).getValue();
    if (value == "") {
      continue;
    }

    try {
      var childFolder = DriveApp.getFolderById(value);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

    try {
      var sharingPermission = childFolder.getSharingPermission();
      ss.getRange(x, lastcolumn + 1).setValue(sharingPermission);
    } catch (e) {
      ss.getRange(x, lastcolumn + 1).setValue("Fail");
      continue;
    }

  }
}