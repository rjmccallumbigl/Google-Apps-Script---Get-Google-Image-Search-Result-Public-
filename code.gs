/******************************************************************************************************
 * Connect to Google Image Search via API, return first result to be placed in cell to the right.
 * 
 * @param {String} query The search value we are searching using the Google Custom Search Engine.
 * @return {String} The URL of the image returned from the custom search.
 * 
 * Instructions
 * 1. Get your Custom Search JSON API key and add it below to var apikey: https://developers.google.com/custom-search/v1/overview#api_key
 * 2. Create search engine, point it to google.com: https://cse.google.com/all
 * 3. In the settings, tell it to enable Image Search, remove any Sites to search, and Search the Entire Web.
 * 4. Copy the search engine ID and add it below to var searchEngineID.
 * 5. Run the script onOpen() and refresh the Google Sheet.
 * 6. Highlight your query in the spreadsheet.
 * 7. Run the script 'Function: Get Google Image Search Result(s)' from the new Functions menu on the Google Sheet.
 * 
 * Sources
 * https://stackoverflow.com/questions/34035422/google-image-search-says-api-no-longer-available
 * https://webmasters.stackexchange.com/questions/18704/return-first-image-source-from-google-images
 * 
 ******************************************************************************************************/
function getGoogleImageSearchResult(query) {

  // Work for several cells
  if (query.map) {
    return query.map(getGoogleImageSearchResult);
  } else {

    // Declare variables
    var numberOfResults = 1;
    var searchType = "image";

    // Add API credentials
    var apikey = "ADD-HERE";
    var searchEngineID = "ADD-HERE";

    // Building call to API
    var url = "https://www.googleapis.com/customsearch/v1?key=" + apikey + "&cx=" + searchEngineID
      + "&q=" + query + "&num=" + numberOfResults + "&searchType=" + searchType;
    console.log(url);

    var params = {
      method: "GET"
    };

    // Calling API
    var response = UrlFetchApp.fetch(url, params);

    // Parsing response
    var responseText = JSON.parse(response.getContentText());
    return responseText.items[0].link;
  }
}
/******************************************************************************************************
 *
 * Return the first Google Image Search image for the item(s) in the selected cell(s) and place them in
 * the cell to the right.
 * 
 ******************************************************************************************************/

function runQuery() {

  // Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getDisplayValues();
  var returnArray = [];
  var returnRange = sheet.getRange(range.getRow(), range.getColumn() + 1, values.length, 1);

  for (var x = 0; x < values.length; x++) {
    returnArray.push(['=IMAGE("' + getGoogleImageSearchResult(values[x]) + '")']);
  }

  // Update sheet
  returnRange.setValues(returnArray);
  SpreadsheetApp.flush();
  returnRange.copyTo(returnRange, { contentsOnly: true });
}

/******************************************************************************************************
* 
* Create a menu option for script functions
*
******************************************************************************************************/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
    .addItem('Function: Get Google Image Search Result(s)', 'runQuery')
    .addToUi();
}
