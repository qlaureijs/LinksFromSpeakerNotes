function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Copy Hyperlinks to Spreadsheet', 'copyHyperlinksToSpreadsheet')
      .addToUi();
}

function copyHyperlinksToSpreadsheet() {
  var presentation = SlidesApp.getActivePresentation();
  var presentationName = presentation.getName();
  var slides = presentation.getSlides();
  
  // Prompt the user to select an existing spreadsheet or create a new one
  var spreadsheetUrl = selectOrCreateSpreadsheet(presentationName);
  if (!spreadsheetUrl) return; // User canceled the operation
  
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheet = spreadsheet.getActiveSheet();
  sheet.clear(); // Clear existing data
  
  // Add header row with swapped columns
  sheet.appendRow(['Slide Number', 'Description', 'URL']);
  
  for (var i = 0; i < slides.length; i++) {
    var slideNumber = i + 1;
    var notes = slides[i].getNotesPage().getSpeakerNotesShape();
    if (notes) {
      var text = notes.getText().asString();
      var matches = text.matchAll(/(?:https?|ftp):\/\/[^\s]+/g); // Match URLs
      
      for (var match of matches) {
        var url = match[0];
        var metadata = fetchMetadata(url);
        
        // Write the hyperlink URL, slide number, and description to the spreadsheet
        sheet.appendRow([slideNumber, metadata.description, url]);
      }
    }
  }
}

function fetchMetadata(url) {
  var metadata = {
    description: ''
  };
  
  try {
    var response = UrlFetchApp.fetch(url);
    var html = response.getContentText();
    
    // Extract description
    var descriptionMatch = html.match(/<meta\s+name=["']description["']\s+content=["'](.*?)["']/i);
    if (descriptionMatch && descriptionMatch.length >= 2) {
      metadata.description = descriptionMatch[1];
    }
  } catch (error) {
    Logger.log('Error fetching metadata for URL: ' + url);
    Logger.log(error);
  }
  
  return metadata;
}

function selectOrCreateSpreadsheet(presentationName) {
  var ui = SlidesApp.getUi();
  var result = ui.prompt(
      'Select or Create Spreadsheet',
      'Enter the URL of an existing Google Spreadsheet or leave it blank to create a new one:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var spreadsheetUrl = result.getResponseText();
  
  if (button == ui.Button.OK) {
    if (spreadsheetUrl.trim() === '') {
      // Create a new spreadsheet with presentation name
      var spreadsheet = SpreadsheetApp.create('Hyperlinks Spreadsheet - ' + presentationName);
      spreadsheetUrl = spreadsheet.getUrl();
    }
    return spreadsheetUrl;
  } else {
    return ''; // User canceled the operation
  }
}
