function yesterdayDate() {
  var d = new Date();
  d.setDate(d.getDate() - 200);
  var formattedWeekAgo = Utilities.formatDate(d, 'GMT', 'MM/dd/yyyy');
  return formattedWeekAgo;
}

function todayDate() {
  var e = new Date();
  var formattedToday = Utilities.formatDate(e, 'GMT', 'MM/dd/yyyy');
  return formattedToday;
}

function primaryContactEntry(contract, index) {
  var primaryContactData = contract["opportunitiesData"][index]["pointOfContact"][0];
  if (primaryContactData.hasOwnProperty("phone") && primaryContactData.hasOwnProperty("email") && primaryContactData.hasOwnProperty("fullName")) {
    var primaryName = contract["opportunitiesData"][index]["pointOfContact"][0]["fullName"];
    var primaryEmail = contract["opportunitiesData"][index]["pointOfContact"][0]["email"];
    var primaryPhone = contract["opportunitiesData"][index]["pointOfContact"][0]["phone"];

    var primaryContact = primaryName + ", " + primaryEmail + ", " + primaryPhone;
    return primaryContact;

  } else if (primaryContactData.hasOwnProperty("email") && primaryContactData.hasOwnProperty("fullName")){
    var primaryEmail = contract["opportunitiesData"][index]["pointOfContact"][0]["email"];
    var primaryPhone = contract["opportunitiesData"][index]["pointOfContact"][0]["phone"];

    var primaryContact = primaryName + ", " + primaryEmail;
    return primaryContact;

  } else {
    var primaryName = contract["opportunitiesData"][index]["pointOfContact"][0]["fullName"];

    var primaryContact = primaryName;
    return primaryContact;
  }
}

function searchString(string){
  var sheet = SpreadsheetApp.getActiveSheet();
  var search_string = string;
  var textFinder = sheet.createTextFinder(search_string);
  return textFinder.findNext();
}

var spreadsheetId = "1ZQL6fkhFKp7Q3d6UefS5De3teEFudyFcn9-98gIkxiQ";
var sheetName = "Main";

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

    // Init of Search Amount
    //var naicsCode = Number(input.naics);
    var searchLimit = 1;

    var solNum = sheet.getRange(newRow, 1).getValue();

    var url = "https://api.sam.gov/prod/opportunities/v2/search?limit=" + searchLimit + "&api_key=TWm9XEEQTCJxPypyTxhUkFGJuwtQDRMVVKpWgMrX&postedFrom=" + yesterday +"&postedTo=" + today + "&solnum=" + solNum;
    var response = UrlFetchApp.fetch(url);
    var contract = JSON.parse(response.getContentText());

    console.log(contract)

    try {
      var title = contract["opportunitiesData"][0]["title"];
      var link = contract["opportunitiesData"][0]["uiLink"];
      sheet.getRange(newRow, 2).setFormula('=HYPERLINK(' + '"' + link + '"' + ', "' + title + '")');
    } catch {
    }

    try {
      var naics = contract["opportunitiesData"][0]["naicsCode"];
      sheet.getRange(newRow, 3).setValue(naics);
    } catch {
    }

    try {
      var deadline = contract["opportunitiesData"][0]["responseDeadLine"];
      deadline = deadline.substr(0, 10);
      sheet.getRange(newRow, 7).setValue(deadline);
    } catch {
    }

    try {
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
