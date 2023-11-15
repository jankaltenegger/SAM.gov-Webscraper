function yesterdayDate() { // Calculates yesterday's date and formats it to MM/dd/yyyy.
  var d = new Date();
  d.setDate(d.getDate() - 200);
  var formattedWeekAgo = Utilities.formatDate(d, 'GMT', 'MM/dd/yyyy');
  return formattedWeekAgo;
}

function todayDate() { // Gets today's date and formats it to MM/dd/yyyy.
  var e = new Date();
  var formattedToday = Utilities.formatDate(e, 'GMT', 'MM/dd/yyyy');
  return formattedToday;
}

// Gets all the necessary contact information of the primary contract of a government solicitation post. 
function primaryContactEntry(contract, index) {
  // Navigates to the first listed contact and stores all the necessary information.
  var primaryContactData = contract["opportunitiesData"][index]["pointOfContact"][0];
  // Case if contact has a name, email and phone number listed.
  if (primaryContactData.hasOwnProperty("phone") && primaryContactData.hasOwnProperty("email") && primaryContactData.hasOwnProperty("fullName")) {
    var primaryName = contract["opportunitiesData"][index]["pointOfContact"][0]["fullName"];
    var primaryEmail = contract["opportunitiesData"][index]["pointOfContact"][0]["email"];
    var primaryPhone = contract["opportunitiesData"][index]["pointOfContact"][0]["phone"];

    var primaryContact = primaryName + ", " + primaryEmail + ", " + primaryPhone;
    return primaryContact;
  // Case if contact has only an email and name listed.
  } else if (primaryContactData.hasOwnProperty("email") && primaryContactData.hasOwnProperty("fullName")){
    var primaryEmail = contract["opportunitiesData"][index]["pointOfContact"][0]["email"];
    var primaryPhone = contract["opportunitiesData"][index]["pointOfContact"][0]["phone"];

    var primaryContact = primaryName + ", " + primaryEmail;
    return primaryContact;
  // If the contact doesn't have a phone number or email listed then gets the name.
  } else {
    var primaryName = contract["opportunitiesData"][index]["pointOfContact"][0]["fullName"];

    var primaryContact = primaryName;
    return primaryContact;
  }
}

// Variables for the wanted Spreadsheet and Sheet
var spreadsheetId = "INSERT SPREADSHEETID HERE";
var sheetName = "NAME";

function main(e) {
  var app = SpreadsheetApp.openById(spreadsheetId);
  var sheet = app.getSheetByName(sheetName);

  var range = e.range;
  var editedCol = range.getColumn();
  var editedSheet = range.getSheet();

  var minCol = 1;
  var maxCol = 1;

  if (editedSheet.getName() === sheetName && editedCol >= minCol && editedCol <= maxCol) {
    const today = todayDate();
    const yesterday = yesterdayDate();

    var newRow = sheet.getLastRow()
    var searchLimit = 1;

    // Gets the solicitation number inserted into the left-most column.
    var solNum = sheet.getRange(newRow, 1).getValue();

    //API call with the desired date range 
    var url = "https://api.sam.gov/prod/opportunities/v2/search?limit=" + searchLimit + "&api_key=TWm9XEEQTCJxPypyTxhUkFGJuwtQDRMVVKpWgMrX&postedFrom=" + yesterday +"&postedTo=" + today + "&solnum=" + solNum;
    var response = UrlFetchApp.fetch(url);
    var contract = JSON.parse(response.getContentText());

    // Following section gets the wanted data and inserts it into the sheet. Should the navigation process throw any exceptions, it means that the required
    // information is not available and the government contract has incomplete/useless information, therefore we skip those parts.
    try {
      // Posts the link of the contract while overlaying it with the name of the contract.
      var title = contract["opportunitiesData"][0]["title"];
      var link = contract["opportunitiesData"][0]["uiLink"];
      sheet.getRange(newRow, 2).setFormula('=HYPERLINK(' + '"' + link + '"' + ', "' + title + '")');
    } catch {
    }

    try {
      // Posts NAICS code
      var naics = contract["opportunitiesData"][0]["naicsCode"];
      sheet.getRange(newRow, 3).setValue(naics);
    } catch {
    }

    try {
      // Posts the deadline, where the government solicitation will either get extended or move onto the next stage.
      var deadline = contract["opportunitiesData"][0]["responseDeadLine"];
      deadline = deadline.substr(0, 10);
      sheet.getRange(newRow, 7).setValue(deadline);
    } catch {
    }

    try {
      // Posts the type of contract it is
      var type = contract["opportunitiesData"][0]["type"];
      sheet.getRange(newRow, 9).setValue(type);
    } catch {
    }

    try {
      // Primary Contact Entry
      sheet.getRange(newRow, 8).setValue(primaryContactEntry(contract, 0));
    } catch {
    }
  }
}
