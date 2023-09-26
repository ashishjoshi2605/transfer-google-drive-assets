// ****** Global Variables Section ******
let sharedDriveFolderIds = new Set();
let sharedDriveNames = new Set();
let sharedDrives = Drive.Drives.list().items;
let foldersNotOnSharedDrive = new Set();
let globalVistedLinks = new Set()
let modifiedLinks = {}
let copiedFileIDMap = {}
let storedLinks = {}
let failedLinks = new Set()
let movedLinks = new Set()
let sharedDriveId;
let rootFilesLink;
let rootFilesType;
var docRegExp = new RegExp("https://docs.google.com/document/d/.*", "");
var driveRegExp = new RegExp("https://drive.google.com/drive(\/u\/[0-9])?/folders/.*", "");
var sheetRegExp = new RegExp("https://docs.google.com/spreadsheets/.*", "");
var formsRegExp = new RegExp("https://docs.google.com/forms/.*", "");
var slidesRegExp = new RegExp("https://docs.google.com/presentation/.*", "");
var folderRegExp = new RegExp("https://drive.google.com/drive(\/u\/[0-9])?/folders/.*", "");
// ****** Global Variables Section ******


/**
* This function manages the flow of the script by calling supporting functions.
*/
function main() {

  // We maintain a separate google spreadsheet to accept user inputs in the form of pop-ups.
  rootFilesLink = takeUserInput();
  if (!rootFilesLink || !sharedDriveId) {
    Logger.log("Error while taking user input....");
    return;
  }

  populateInitialSharedDriveFolders();

  retriveData();

  Logger.log("Root docs Links: %s", rootFilesLink);

  // go through each root file. In case of Google doc, parse the document to find more links.
  // rootFilesLink.forEach(fileUrl => {
  try {
    OpenFilesByUrlRecursively(rootFilesLink);
  }
  catch (err) {
    Logger.log('Unhandled exception has occurred. Failed with error %s', err.message);
  }
  // });
  Logger.log("Total successfully moved files: %s", Array.from(movedLinks.values()));
  Logger.log("Total successfully copied files: %s", JSON.stringify(modifiedLinks));
  Logger.log("Total move and copy failed files: %s", Array.from(failedLinks.values()));
}

/**
* This function finds the shared drives on the owner is part and stores them in set.
*/
function populateInitialSharedDriveFolders() {
  sharedDrives.forEach((drive) => {
    sharedDriveFolderIds.add(drive.id);
    sharedDriveNames.add(drive.name);
  })
  sharedDriveFolderIds.add(DriveApp.getRootFolder().getId());
  sharedDriveNames.add(DriveApp.getRootFolder().getName());
}

// Function to get current timestamp
function getCurrentTimeStamp() {
  var date = new Date();
  var estDate = Utilities.formatDate(date, "America/New_York", "yyyy-MM-dd HH:mm:ss");
  return estDate;
}

/**
* This function finds retrieves the previously moved and copied files links along with sharedDriveId
* from the container spreadsheet and stores them in a globally declared dictionary.
*/
function retriveData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Dictionary");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] != "FAILED") {
      storedLinks[data[i][1]] = data[i][2];
    }
  }
}

/**
* This function finds appends the row passed to it in the container spreadsheet.
* and stores them in a globally declared dictionary.
* @param {List} row to be appended to the spreadsheet
*/
function saveData(row, sheetName) {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // If sheet doesn't exist, create a new sheet
  if (!sheet) {
    spreadsheet.insertSheet(sheetName);
    sheet = spreadsheet.getSheetByName(sheetName);
  }
  sheet.appendRow(row);
}




/**
* This function accepts input (through container spreadsheet popups) for root file and destination shared drive id.
* The root file can be Google Spreadsheet, Google Doc, Google Folder.
* @return {Set} set containing rootFileLinks
*/
function takeUserInput() {
  // var rootFilesLinksSet = new Set();
  var ui = SpreadsheetApp.getUi();

  var rootFileInput = ui.prompt(
    'Please enter the URL of root Google Folder or Google Spreadsheet or Google Doc url to traverse:',
    ui.ButtonSet.OK_CANCEL);

  // Prints out an error message in the event that Spreadsheet URL not provided.
  if (rootFileInput.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log("We cannot move ahead without Google Spreadsheet or Google Doc Link. Exiting script!");
    return;
  }

  var rootFileInputLink = rootFileInput.getResponseText();
  Logger.log("Root Google Link: %s", rootFileInputLink);

  // decode url if it is encoded.
  rootFileInputLink = decodeUrl(rootFileInputLink);

  // process url based on url type
  let docsLink = docRegExp.exec(rootFileInputLink);
  let sheetsLink = sheetRegExp.exec(rootFileInputLink);
  let folderLink = folderRegExp.exec(rootFileInputLink);
  let finalFileLink = "";
  rootFilesType = "";
  if (docsLink) {
    finalFileLink = docsLink[0];
    rootFilesType = "doc";
    finalFileLink = appendEditTextToUrl(finalFileLink)
    // rootFilesLinksSet.add(appendEditTextToUrl(finalFileLink));
  }
  else if (sheetsLink) {
    finalFileLink = sheetsLink[0];
    rootFilesType = "sheet";
    // rootFilesLinksSet = extractLinksFromSheet1(finalFileLink);
  }
  else if (folderLink) {
    finalFileLink = folderLink[0];
    rootFilesType = "folder";
    // var rootFolderId = getIdFromUrl(finalFileLink)
    // Logger.log("Root folder id: %s", rootFolderId);
    // getFilesFromFolder(rootFolderId, rootFilesLinksSet);
  }
  else {
    Logger.log("Invalid Root Google Docs Link");
    return;
  }

  // Accepts user input for the destination shared Drive id. 
  // This is where the user expects to move/copy google document files into for future reference.
  var sharedDriveURLInput = ui.prompt(
    'Please enter the destination Shared Drive URL:',
    ui.ButtonSet.OK_CANCEL);

  // Prints out an error message in the event that shared id input is not provided.
  if (sharedDriveURLInput.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log("We cannot move ahead without destination Shared Drive URL. Exiting script!");
    return;
  }

  sharedDriveURLInput = driveRegExp.exec(sharedDriveURLInput.getResponseText());
  Logger.log("Shared Drive input url: %s",sharedDriveURLInput);
  if (sharedDriveURLInput) {
    let sharedDriveURL = sharedDriveURLInput[0];
    sharedDriveId = getIdFromUrl(sharedDriveURL)
    Logger.log("Shared Drive URL id: %s", sharedDriveId);
  } else {
      return;
  }

  return finalFileLink;
}

/**
 * https://gist.github.com/mogsdad/6518632
 * Get an array of all LinkUrls in the document. The function is
 * recursive, and if no element is provided, it will default to
 * the active document's Body element.
 *
 * @param {Element} element The document element to operate on. 
 * .
 * @returns {Array}         Array of objects, vis
 *                              {element,
 *                               startOffset,
 *                               endOffsetInclusive, 
 *                               url}
 */

function getAllLinks(element, documentURL, folderLinkSet) {
  var links = [];
  element = element || DocumentApp.getActiveDocument().getBody();
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var textObj = element.editAsText();
    var text = element.getText();
    for (var ch = 0; ch < text.length; ch++) {
      var url = textObj.getLinkUrl(ch);
      var docsLink = docRegExp.exec(url);
      var sheetsLink = sheetRegExp.exec(url);
      var formLink = formsRegExp.exec(url);
      var slidesLink = slidesRegExp.exec(url);
      var folderLink = folderRegExp.exec(url);
      if(folderLink && !folderLinkSet.has(url))
      {
        folderLinkSet.add(url);
        var documentName = DriveApp.getFileById(getFileIdFromUrl(documentURL)).getName();
        var folderName = DriveApp.getFolderById(getIdFromUrl(url)).getName();
        var documentLinkFormula = '=HYPERLINK("' + documentURL + '", "' + documentName + '")';
        var folderFormula = '=HYPERLINK("' + url + '", "' + folderName + '")';
        saveData([documentLinkFormula, folderFormula], "FolderLinks");
      }
      if (docsLink || sheetsLink || formLink || slidesLink) {
        url = decodeUrl(url);
        if (docsLink) {
          url = docsLink[0];
        }
        else if (formLink) {
          url = formLink[0];
        }
        else if (slidesLink) {
          url = slidesLink[0];
        }
        else {
          url = sheetsLink[0];
          // let extracted_links = extractLinksFromSheet1(url);
          // extracted_links.forEach((extracted_link) => {
          //   var curUrl = {};
          //   curUrl.element = element;
          //   curUrl.url = String(extracted_link); // grab a copy
          //   curUrl.startOffset = ch;
          //   curUrl.endOffsetInclusive = ch + extracted_link.length - 1;
          //   ch = ch + extracted_link.length;
          //   links.push(curUrl);  // add to links
          //   curUrl = {};
          // });
        }
        let lindex = url.lastIndexOf("/edit");
        if (lindex != -1) {
          url = url.substring(0, lindex + 5);
        }
        else {
          if (url[url.length - 1] == '/') {
            url += "edit"
          }
          else {
            url += "/edit"
          }
        }
        var curUrl = {};
        curUrl.element = element;
        curUrl.url = String(url); // grab a copy
        curUrl.startOffset = ch;
        curUrl.endOffsetInclusive = ch + url.length - 1;
        ch = ch + url.length;
        links.push(curUrl);  // add to links
        curUrl = {};
      }
    }
  }
  else if (element.getType() === DocumentApp.ElementType.RICH_LINK) {
    var curUrl = {};
    curUrl.element = element;
    curUrl.url = element.getUrl(); // grab a copy
    var docsLink = docRegExp.exec(curUrl.url);
    var sheetsLink = sheetRegExp.exec(curUrl.url);
    var formLink = formsRegExp.exec(curUrl.url);
    var slidesLink = slidesRegExp.exec(curUrl.url);
    var folderLink = folderRegExp.exec(curUrl.url);
    if (docsLink || sheetsLink || formLink || slidesLink) {
      links.push(curUrl.url);
    }
    if(folderLink && !folderLinkSet.has(curUrl.url))
    {
      folderLinkSet.add(curUrl.url);
      var documentName = DriveApp.getFileById(getFileIdFromUrl(DocumentApp.getActiveDocument().getUrl())).getName();
      var folderName = DriveApp.getFolderById(getIdFromUrl(curUrl.url)).getName();
      var documentLinkFormula = '=HYPERLINK("' + DocumentApp.getActiveDocument().getUrl() + '", "' + documentName + '")';
      var folderFormula = '=HYPERLINK("' + curUrl.url + '", "' + folderName + '")';
      saveData([documentLinkFormula, folderFormula], "FolderLinks");
    }
  }
  else {
    // Get number of child elements, for elements that can have child elements. 
    try {
      var numChildren = element.getNumChildren();
    }
    catch (e) {
      numChildren = 0;
    }
    for (var i = 0; i < numChildren; i++) {
      links = links.concat(getAllLinks(element.getChild(i), documentURL, folderLinkSet));
    }
  }
  return links;
}

/**
 * Work function. Open the Google Doc by the given URL. Add the current url to the global visited set.
 * @param url
 * @return {curr_links} A set including all url found (NEED TO REFACTOR THIS PART IF THE # OF DOCUMENTS IS LARGE.)
 */

// Maintain a queue of documents encountered but not yet traversed.  Call it docsToProcess.
// As you scan a document, when you encounter a link, add it to the end of the queue.
// After you finish scanning a document, take the head entry from docsToProcess, and scan it for links.
// Read Jones & Lins, Garbage Collection: Algorithms for Dynamic Memory Management
function OpenFilesByUrlRecursively(url) {
  url = decodeUrl(url)
  if (globalVistedLinks.has(url)) {
    Logger.log('Found cycle with url %s', url);
    return;
  }
  
  try {
    globalVistedLinks.add(url);
    let docsLinkCheck = docRegExp.exec(url);
    let sheetsLinkCheck = sheetRegExp.exec(url);
    let folderLinkCheck = folderRegExp.exec(url);
    let copiedDocument = false;

    var curr_links = [];
    var sheet_links = [];
    var folderFilesLinksSet = new Set();
    var curr_body;

    try {
      if((url != rootFilesLink) || ((url == rootFilesLink) && (rootFilesType != "folder"))){
        d_index = url.lastIndexOf("/d/");
        edit_index = url.lastIndexOf("/edit");
        file_id = url.substring(d_index + 3, edit_index);
        if (storedLinks.hasOwnProperty(url) && storedLinks[url] == sharedDriveId) {
          Logger.log('URL visited in previous execution of script: %s', url);
          saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, url, sharedDriveId, 'VISITED'], "Logs");
        }
        //  else if (checkIfFileOnSharedDrive(file_id)) {
        //   Logger.log('File already present on shared drive: %s', url);
        //  }
         else if (checkOwnerAndExistence(file_id)) {
           Logger.log('File owned by user and present in user\'s drive', url);
           saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, url, sharedDriveId, 'EXISTS'], "Logs");
         }
        else if (!moveToSharedDrive(file_id, sharedDriveId)) {
          let trimUrl = url.substring(0,edit_index)+"/edit";
          if (!modifiedLinks[getFileIdFromUrl(url)] && !modifiedLinks[getFileIdFromUrl(trimUrl)]) {
            var copiedFileUrl = copyToSharedDrive(file_id, sharedDriveId);
            copiedDocument = true;
            modifiedLinks[getFileIdFromUrl(url)] = copiedFileUrl;
            saveData([url, modifiedLinks[getFileIdFromUrl(url)], sharedDriveId, 'COPIED'], "Dictionary");
            saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, modifiedLinks[getFileIdFromUrl(url)], sharedDriveId, 'COPIED'], "Logs");
          }
        }
        else {
          movedLinks.add(url);
          saveData([url, url, sharedDriveId, 'MOVED'], "Dictionary");
          saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, url, sharedDriveId, 'MOVED'], "Logs");
        }
        Logger.log("Done with %s", url);
      }
    }
    catch (err) {
      Logger.log(err.message);
      Logger.log("Failed for file %s \n continuing with next file.", url);
      failedLinks.add(url);
      saveData([url, url, sharedDriveId, 'FAILED'], "Dictionary");
      saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, url, sharedDriveId, 'FAILED'], "Logs");
    }

    // Extracting links in the current file if the file is a google doc
    if (docsLinkCheck) { 
      var doc = DocumentApp.openByUrl(url);
      var curr_body = doc.getBody();
      
      if(modifiedLinks[getFileIdFromUrl(url)])
      {
        curr_links = getAllLinks(curr_body, modifiedLinks[getFileIdFromUrl(url)], new Set());
      }
      else
      {
        curr_links = getAllLinks(curr_body, url, new Set());
      }
      Logger.log('Document URL: %s', url);
      Logger.log('Document links are: %s', curr_links);
      curr_links.forEach(element => OpenFilesByUrlRecursively(element["url"] || element));
    }
    // Extracting links in the current file if the file is a google sheet
    else if(sheetsLinkCheck) {
      sheet_links = extractLinksFromSheet1(url);
      Logger.log('Sheet links are: %s', sheet_links);
      sheet_links.forEach(link => OpenFilesByUrlRecursively(link));
    }
    // Extracting links of files in the current folder
    else if(folderLinkCheck) {
      var folderId = getIdFromUrl(url);
      Logger.log("Folder id: %s", folderId);
      getFilesFromFolder(folderId, folderFilesLinksSet);
      folderFilesLinksSet.forEach(file => OpenFilesByUrlRecursively(file));
    }
    
    if(docsLinkCheck) {
      if (!copiedDocument) {
        Logger.log("Updating doc urls inside %s.", url);
        updateCopiedDocLinks(curr_links, url);
      } else {
        Logger.log("Updating doc urls inside new url %s.", modifiedLinks[getFileIdFromUrl(url)]);
        updateCopiedDocLinks(curr_links, modifiedLinks[getFileIdFromUrl(url)]);
      }
      saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, url, sharedDriveId, 'READ COMPLETE'], "Logs");
    }
    if(sheetsLinkCheck) {
      Logger.log("Updating sheet urls inside %s.", url);
      updateCopiedSheetLinks(url);
      saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, url, sharedDriveId, 'READ COMPLETE'], "Logs");
    }
  }
  catch (err) {
    Logger.log('Failed with error %s %s', err.message, url);
    failedLinks.add(url);
    saveData([url, url, sharedDriveId, 'FAILED'], "Dictionary");
    saveData([getCurrentTimeStamp(), getFileNameFromUrl(url), url, url, sharedDriveId, 'FAILED'], "Logs");
  }
}

/*
* This function updates the original URLs by their copied URLs in the original spreadsheet
* @param {sheet_url} the original spreadsheet where the URLs need to be updated
*/
function updateCopiedSheetLinks(sheet_url){
  if(modifiedLinks[getFileIdFromUrl(sheet_url)])
    sheet_url = modifiedLinks[getFileIdFromUrl(sheet_url)];
  let range = SpreadsheetApp.openByUrl(sheet_url).getDataRange();
  let nRows = range.getNumRows();
  let nCols = range.getNumColumns();
  for(var i=1; i<= nRows; i++){
    for(var j=1; j<= nCols; j++){
      let cell = range.getCell(i,j);
      let currVal = cell.getValue();
      if (currVal == null)// || typeof currVal != 'string')
        continue;
      if (typeof currVal != 'string')
          currVal = currVal.toString();
      currVal = decodeUrl(currVal);
      if (currVal == null)
        continue;
      currVal = appendEditTextToUrl(currVal);
      if(currVal && modifiedLinks[getFileIdFromUrl(currVal)]){
        cell.setValue(modifiedLinks[getFileIdFromUrl(currVal)]);
      }
    }
  }
  
  var richTextValues = range.getRichTextValues();
  for (var i = 0; i < richTextValues.length; i++) {
    for (var j = 0; j < richTextValues[0].length; j++) {
      var value = richTextValues[i][j];
      if (value && value.getLinkUrl()) {
        var retreivedLink = decodeUrl(value.getLinkUrl());
        var docsLink = docRegExp.exec(retreivedLink);
        var sheetsLink = sheetRegExp.exec(retreivedLink);
        var formLink = formsRegExp.exec(retreivedLink);
        var slidesLink = slidesRegExp.exec(retreivedLink);

        if (docsLink || sheetsLink || formLink || slidesLink) {
          if(modifiedLinks[getFileIdFromUrl(retreivedLink)]){
            const new_value = SpreadsheetApp.newRichTextValue().setText(value.getText()).setLinkUrl(modifiedLinks[getFileIdFromUrl(retreivedLink)])
                  .build()
            richTextValues[i][j] = new_value;
          }
        }
      } 
      else if(!value.getText()) {
        var cellValue = range.getCell(i+1,j+1).getValue().toString();
         richTextValues[i][j] = SpreadsheetApp.newRichTextValue().setText(cellValue).build();
      }
    }
  }
  range.setRichTextValues(richTextValues);
}

function replaceHyperlinks(sheet_url, ) {
  if(modifiedLinks[getFileIdFromUrl(sheet_url)])
    sheet_url = modifiedLinks[getFileIdFromUrl(sheet_url)];
  var sheet = SpreadsheetApp.openByUrl(sheet_url)
  var range = sheet.getDataRange();
  var formulas = range.getFormulas();
  
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[0].length; j++) {
      var formula = formulas[i][j];
      if (formula.startsWith('=HYPERLINK(')) {
        var parts = formula.match(/"(.*?)"/g);
        var url = parts[0].replace(/"/g, ""); // remove the quotes from the URL
        var label = parts[1].replace(/"/g, ""); // remove the quotes from the label
        var newUrl = "https://www.example.com"; // replace with your new URL
        
        var newFormula = '=HYPERLINK("' + newUrl + '","' + label + '")'; // create the new hyperlink formula
        sheet.getRange(i+1, j+1).setFormula(newFormula); // set the new formula in the cell
      }
    }
  }
}

function updateCopiedDocLinks(curr_links, curr_url) {
    curr_body = DocumentApp.openByUrl(curr_url).getBody();
    try {
    curr_links.forEach(element => {   
    let link = element["url"] || element;
    if (modifiedLinks[getFileIdFromUrl(link)]) {
      let replaceString = modifiedLinks[getFileIdFromUrl(link)];

      let file_id = getFileIdFromUrl(link);

      let regex_pattern = "^.*" + file_id + ".*$"; 
      let foundText = curr_body.findText(regex_pattern);
      if (foundText != null) {
        let startText = foundText.getStartOffset();
        let endText = startText + replaceString.length - 1;
        let element = foundText.getElement();
        element.asText()
          .replaceText(regex_pattern, replaceString)
          .setLinkUrl(startText, endText, replaceString);
      }
      
      var textObj = curr_body.editAsText();
      var text = curr_body.getText();

      for (var ch = 0; ch < text.length; ch++) {
        var url = textObj.getLinkUrl(ch);
        if(url == null)
        {
          continue;
        }
        let url_file_id = getFileIdFromUrl(url);
        let link_file_id = getFileIdFromUrl(link);
        if(url && url_file_id == link_file_id) {
          url = decodeUrl(url);
          let lindex = url.lastIndexOf("/edit");
          if (lindex != -1) {
            url = url.substring(0, lindex + 5);
          }
          else {
            if (url[url.length - 1] == '/') {
              url += "edit"
            }
            else {
              url += "/edit"
            }
          }
          let count = 1;
          for(var ch2= ch+1; ch2<text.length; ch2++){
            var l = textObj.getLinkUrl(ch2);
            if (l == null)
            {
              continue;
            }
            let l_file_id = getFileIdFromUrl(l);
            if(l && l_file_id == link_file_id)
              count++;
            else
              break;
          }
          textObj.setLinkUrl(ch,ch + count - 1, modifiedLinks[getFileIdFromUrl(link)]);
          ch = ch + count - 1;
        }
      }

      replaceRichLinks(curr_url);
      
    }
  })
    } catch (err) {
    Logger.log("Error while updating  copied doc links %s", err);
  }
}

function replaceRichLinks(curr_url) {
  // Get the active Google Doc
  var doc = DocumentApp.openByUrl(curr_url);

  // Get the body of the document
  var body = doc.getBody();

  // Get all the paragraphs in the body
  var paragraphs = body.getParagraphs();

  // Loop through each paragraph and remove the rich link elements
  for (var i = 0; i < paragraphs.length; i++) {
    var paragraph = paragraphs[i];

    // Get all the elements in the paragraph
    var elements = paragraph.getNumChildren();

    // Loop through each element in the paragraph
    for (var j = 0; j < elements; j++) {
      var element = paragraph.getChild(j);

      // Check if the element is a rich link
      if (element.getType() == DocumentApp.ElementType.RICH_LINK) {

        var textUrl = element.asRichLink().getUrl();
        // Replace the rich link element with a text element containing the same text
        paragraph.insertText(j, modifiedLinks[getFileIdFromUrl(textUrl)]);
        paragraph.removeChild(element);
      }
    }
  }
}






/**
* This function moves the file to the globally declared shared drive.
* @param {string} id of file to be moved
* @return {bool} if file moved successfully or not
*/
function moveToSharedDrive(srcFileId) {
  try {
    Logger.log("File Id - %s", srcFileId);
    var srcFile = DriveApp.getFileById(srcFileId);
    var destination = DriveApp.getFolderById(sharedDriveId);
    srcFile.moveTo(destination);
    Logger.log("Successfully moved %s", srcFile.getName());
    return true;
  } catch (err) {
    Logger.log("Error %s", err);
    return false;
  }
}

/**
* This function copies the file to the globally declared shared drive.
* @param {string} id of file to be copied
* @return {url} url of copy of file
*/
function copyToSharedDrive(srcFileId) {
  var srcFile = DriveApp.getFileById(srcFileId);
  var destination = DriveApp.getFolderById(sharedDriveId);
  var copiedFile = srcFile.makeCopy(srcFile.getName(), destination);
  Logger.log("Successfully copied %s, new url is %s", srcFile.getName(), copiedFile.getUrl());
  return copiedFile.getUrl();
}

/**
* This function parses the spreadsheet and extracts all the existing URLs in it. Extraction using regex.
* @param {string} sheet url
* @return {Set} extracted links from sheet
*/
function extractLinksFromSheet1(sheetUrl) {
  var sheet = SpreadsheetApp.openByUrl(sheetUrl);
  var finalExtractedLinks = new Set();
  var data = sheet.getDataRange().getValues();
  var folderLinkSet = new Set();
  var modifiedSheetUrl;
  if(modifiedLinks[getFileIdFromUrl(sheetUrl)])
  {
    modifiedSheetUrl = modifiedLinks[getFileIdFromUrl(sheetUrl)];
  }
  else
  {
    modifiedSheetUrl = sheetUrl;
  }
  data.forEach(function (row) {
    row.forEach(function (retreivedLink) {
      
      retreivedLink = decodeUrl(retreivedLink);

      // process link based on it's type
      var docsLink = docRegExp.exec(retreivedLink);
      var sheetsLink = sheetRegExp.exec(retreivedLink);
      var formLink = formsRegExp.exec(retreivedLink);
      var slidesLink = slidesRegExp.exec(retreivedLink);
      var folderLink = folderRegExp.exec(retreivedLink);

      if (folderLink && !folderLinkSet.has(retreivedLink))
      {
        folderLinkSet.add(retreivedLink);
        var documentName = DriveApp.getFileById(getFileIdFromUrl(modifiedSheetUrl)).getName();
        var folderName = DriveApp.getFolderById(getIdFromUrl(retreivedLink)).getName();
        var documentLinkFormula = '=HYPERLINK("' + modifiedSheetUrl + '", "' + documentName + '")';
        var folderFormula = '=HYPERLINK("' + retreivedLink + '", "' + folderName + '")';
        saveData([documentLinkFormula, folderFormula], "FolderLinks");
      }

      if (docsLink || sheetsLink || formLink || slidesLink) {
        console.log("retreivedLink : %s", retreivedLink);
        finalExtractedLinks.add(appendEditTextToUrl(retreivedLink));
      }
    });
  });
  var range = sheet.getDataRange();
  var richTextValues = range.getRichTextValues();
  
  for (var i = 0; i < richTextValues.length; i++) {
    for (var j = 0; j < richTextValues[0].length; j++) {
      var value = richTextValues[i][j];
      if (value && value.getLinkUrl()) {
        retreivedLink = decodeUrl(value.getLinkUrl());
        Logger.log("Hyperlink found: " + retreivedLink);

        var docsLink = docRegExp.exec(retreivedLink);
        var sheetsLink = sheetRegExp.exec(retreivedLink);
        var formLink = formsRegExp.exec(retreivedLink);
        var slidesLink = slidesRegExp.exec(retreivedLink);
        var folderLink = folderRegExp.exec(retreivedLink);
        if(folderLink && !folderLinkSet.has(retreivedLink))
        {
          folderLinkSet.add(retreivedLink);
          var documentName = DriveApp.getFileById(getFileIdFromUrl(modifiedSheetUrl)).getName();
          var folderName = DriveApp.getFolderById(getIdFromUrl(retreivedLink)).getName();
          var documentLinkFormula = '=HYPERLINK("' + modifiedSheetUrl + '", "' + documentName + '")';
          var folderFormula = '=HYPERLINK("' + retreivedLink + '", "' + folderName + '")';
          saveData([documentLinkFormula, folderFormula], "FolderLinks");
        }
        if (docsLink || sheetsLink || formLink || slidesLink) {
          finalExtractedLinks.add(appendEditTextToUrl(retreivedLink));
        }
      }
    }
  }

  return finalExtractedLinks;
}

/**
* This function checks if file is part of user owned shared drives.
* @param {string} File Id
* @return {bool} file part of shared drive or not
*/
function checkIfFileOnSharedDrive(fileId) {
  try {
    let file = DriveApp.getFileById(fileId);
    let parents = file.getParents();
    let tempFolderIds = new Set();
    let check = getParent(parents, tempFolderIds);
    if (check) {
      tempFolderIds.forEach(folderId => {
        sharedDriveFolderIds.add(folderId);
      });
      Logger.log("Files '%s' on shared drive", Array.from(sharedDriveFolderIds));
    }
    else {
      tempFolderIds.forEach(folderId => {
        foldersNotOnSharedDrive.add(folderId);
      });
      Logger.log("Files '%s' not on shared drive", Array.from(foldersNotOnSharedDrive));
    }
    return check;
  }
  catch (err) {
    Logger.log("File with fileID '%s' is not accessible!!", fileId);
    return false;
  }
}

function checkOwnerAndExistence(fileId){
  var check1 = false;
  var check2 = false;
  var file = DriveApp.getFileById(fileId);
  var fileOwner = file.getOwner();
  if (fileOwner.getEmail() === Session.getActiveUser().getEmail()) {
    Logger.log("File is owned by the user");
    check1 = true;
  }

   // Check if the file already exists in the owner's Google Drive
  var fileName = file.getName();
  var files = DriveApp.getFilesByName(fileName);
  while (files.hasNext()) {
    var existingFile = files.next();
    if (existingFile.getId() === fileId) {
      Logger.log("File already exists in the user's Google Drive");
      check2 = true;
    }
  }
  return check1 && check2;
}


function getParent(parents, tempFolderIds) {
  while (parents.hasNext()) {
    let parent = parents.next();
    tempFolderIds.add(parent.getId())
    if (sharedDriveFolderIds.has(parent.getId())) {
      return true;
    }
    else if (foldersNotOnSharedDrive.has(parent.getId())) {
      return false;
    }
    let parentList = parent.getParents();
    return getParent(parentList, tempFolderIds);
  }
  return false;
}

/**
* This function recursively adds all file links to rootFilesLink set recursively(from folders and subfolders).
* @param {string} Folder Id
* @param {Set} rootFilesLinks set that gets populated with file links found in folder
*/
function getFilesFromFolder(folderId, rootFilesLinks) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  while (files.hasNext()) {
    file = files.next();
    rootFilesLinks.add(file.getUrl());
  }
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    getFilesFromFolder(subFolders.next().getId(), rootFilesLinks);
  }
}

/**
* This function decodes encoded url.
* @param {string} Encoded File Url
* @return {string} decoded file url
*/
function decodeUrl(encodedUrl) {
  try{
    if(encodedUrl && encodedUrl.length > 0) {
      encodedUrl = decodeURIComponent(encodedUrl.replace(/\+/g, " "));
      var encIndex = encodedUrl.lastIndexOf("http://www.google.com/url?q=");
      if (encIndex != -1) {
        encodedUrl = encodedUrl.substring(encIndex + 1, url.length);
      }
    }
  }
  catch(err){
    Logger.log("Decoding failed for some text/URL");
  }
  return encodedUrl;
}

/**
* This function appends 'edit' to url if needed.
* @param {string} File Url
* @return {string} Updated File Url
*/
function appendEditTextToUrl(fileUrl) {
  if(fileUrl){
    let lindex = fileUrl.lastIndexOf("/edit");
    if (lindex != -1) {
      fileUrl = fileUrl.substring(0, lindex + 5);
    }
    else {
      if (fileUrl[fileUrl.length - 1] == '/') {
        fileUrl += "edit"
      }
      else {
        fileUrl += "/edit"
      }
    }
  }
  return fileUrl;
}

function getFileIdFromUrl(url) {
  var parts = url.match(/\/d\/(.+)\//);
  if (parts == null || parts.length < 2) {
    return url;
  } else {
    return parts[1];
  }
}

function getFileNameFromUrl(url) {
  var srcFileId = getFileIdFromUrl(url);
  var srcFile = DriveApp.getFileById(srcFileId);
  return srcFile.getName();
}

/**
* This function extracts id from a url.
* @param {string} File Url
* @return {string} File id
*/
function getIdFromUrl(fileUrl) {
    fileUrl = decodeUrl(fileUrl);
    Logger.log("FileUrl inside getIdFromUrl method: ",fileUrl);
    var fileId = fileUrl.substring(fileUrl.lastIndexOf('/') + 1);
    var qIndex = fileId.lastIndexOf('?');
    if (qIndex != -1) {
      fileId = fileId.substring(0, qIndex);
    }
    return fileId;
}
