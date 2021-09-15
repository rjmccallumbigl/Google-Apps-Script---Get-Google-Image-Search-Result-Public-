# Google-Apps-Script---Get-Google-Image-Search-Result-Public-
Connect to Google Image Search via API, return first result to be placed in cell to the right.

* Create a new Google Sheet: https://sheets.new
  * Click Tools -> Script Editor
  * Delete all the text in the script editor and paste the above script
  * Run the function `runQuery()`

* Head to that link and click Get a Key, then you'll have to follow the directions to create a new Google Cloud project: https://developers.google.com/custom-search/v1/overview#api_key
  * click Get a Key
  * Enter a new project name
  * Copy the generated API key
  * Paste it into the GAS code for the variable `apikey`

* Settings: https://cse.google.com/all
  * Add a search engine if you haven't already
  * Site to search: www.google.com
  * When you save it, go to the Control Panel for the search engine (or Edit search engine)
  * enable Image Search (switch it from Off to On)
  * Copy the generated Search engine ID
  * Paste it into the GAS code for the variable `searchEngineID`
  * Under **Sites to search**, select the site entered earlier (should be www.google.com) and delete it (might work without deleting it as well, you can leave it for testing but I just deleted mine it case it was restricted to just google.com)
  * enable Search the Entire Web (switch it from Off to On)

* Try it out!
  * Go back to your Google Sheet and type something you want to search in a cell. 
  * Highlight this typed text. 
  * Click the new menu at the top to the right of Help and select the function: 
   * Functions -> Function: Get Google Image Search Result(s)
  * When you run it, accept the permissions or whatever
  * The first result from your Google Image Search should appear in the cell to the right

![image](https://user-images.githubusercontent.com/15747450/133480269-2430f47b-0fb3-4da8-ad8b-03edc77240e9.png)

