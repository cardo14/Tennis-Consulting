
function copyPointScore() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastColumnIndex = sheet.getLastColumn();
  for (var i = 1; i <= data.length; i++) { // loop through each row
    if (data[i - 1][getColumnIndex("isPointStart", sheet, lastColumnIndex)] == 1) { // if isPointStart is equal to 1
      var pointScore = data[i - 1][getColumnIndex("pointScore", sheet, lastColumnIndex)]; // get the pointScore
      var endIndex = findLastEndIndex(i - 1, data, sheet, lastColumnIndex); // get the index of the last row labeled with isPointEnd
      if (endIndex !== -1) { // if we have an endIndex
        var startIndex = findClosestStartIndex(endIndex, data, sheet, lastColumnIndex); // gets the closest isPointStart
        if (endIndex == (i - 1)) {
          data[endIndex][getColumnIndex("pointScore", sheet, lastColumnIndex)] = pointScore;
          continue;
        }
        for (var j = startIndex; j <= endIndex; j++) { // loop from the top down and copy score down
          data[j][getColumnIndex("pointScore", sheet, lastColumnIndex)] = pointScore;
        }
      }
    }
  }

  sheet.getRange(1, 1, data.length, lastColumnIndex).setValues(data);
}


function copyTiebreakScore() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastColumnIndex = sheet.getLastColumn();
  for (var i = 1; i <= data.length; i++) { // loop through each row
    if (data[i - 1][getColumnIndex("isPointStart", sheet, lastColumnIndex)] == 1) { // if isPointStart is equal to 1
      var tiebreakScore = data[i - 1][getColumnIndex("tiebreakScore", sheet, lastColumnIndex)]; // get the tiebreakScore
      var endIndex = findLastEndIndex(i - 1, data, sheet, lastColumnIndex); // get the index of the last row labeled with isPointEnd
      if (endIndex !== -1) { // if we have an endIndex
        var startIndex = findClosestStartIndex(endIndex, data, sheet, lastColumnIndex); // gets the closest isPointStart
        if (endIndex == (i - 1)) {
          data[endIndex][getColumnIndex("tiebreakScore", sheet, lastColumnIndex)] = tiebreakScore;
          continue;
        }
        for (var j = startIndex; j <= endIndex; j++) { // loop from the top down and copy score down
          data[j][getColumnIndex("tiebreakScore", sheet, lastColumnIndex)] = tiebreakScore;
        }
      }
    }
  }

  sheet.getRange(1, 1, data.length, lastColumnIndex).setValues(data);
}

function copyGameScore() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var data = sheet.getDataRange().getValues();
 var lastColumnIndex = sheet.getLastColumn();
 for (var i = 1; i <= data.length; i++) { // loop through each row
   if (data[i - 1][getColumnIndex("isPointStart", sheet, lastColumnIndex)] == 1) { // if isPointStart is equal to 1
     var gameScore = data[i - 1][getColumnIndex("gameScore", sheet, lastColumnIndex)]; // get the pointScore
     var endIndex = findLastEndIndex(i - 1, data, sheet, lastColumnIndex); // get the index of the last row labeled with isPointEnd
     if (endIndex !== -1) { // if we have an endIndex
       var startIndex = findClosestStartIndex(endIndex, data, sheet, lastColumnIndex); // gets the closest isPointStart
       if (endIndex == (i - 1)) {
         data[endIndex][getColumnIndex("gameScore", sheet, lastColumnIndex)] = gameScore;
         continue;
       }
       for (var j = startIndex; j <= endIndex; j++) { // loop from the top down and copy score down
         data[j][getColumnIndex("gameScore", sheet, lastColumnIndex)] = gameScore;
       }
     }
   }
 }


 sheet.getRange(1, 1, data.length, lastColumnIndex).setValues(data);
}

function updateGameScore() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Set proper score in proper column
 lastRow = sheet.getLastRow()
 var gameScore = sheet.getRange('B' + (lastRow - 1)).getValue();
 sheet.getRange('B' + (lastRow)).setValue(gameScore);

}


// Function to find the index of the last row labeled with isPointEnd
function findLastEndIndex(startIndex, data, sheet, lastColumnIndex) {
  if (data[startIndex][getColumnIndex("isPointEnd", sheet, lastColumnIndex)] == 1) {
    return startIndex;
  }
  for (var i = startIndex + 1; i < data.length; i++) {
    if (data[i][getColumnIndex("isPointEnd", sheet, lastColumnIndex)] == 1) {
      return i;
    }
  }
  return -1; // If no isPointEnd is found
}

// Function to find the closest row labeled with isPointStart
function findClosestStartIndex(endIndex, data, sheet, lastColumnIndex) {
  for (var i = endIndex - 1; i >= 0; i--) {
    if (data[i][getColumnIndex("isPointStart", sheet, lastColumnIndex)] == 1) {
      return i;
    }
  }
  return 0; // If no isPointStart is found, return the first row
}

// Function to get the column index by name
function getColumnIndex(columnName, sheet, lastColumnIndex) {
  var headers = sheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0];
  return headers.indexOf(columnName);
}


function loveAll() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("0-0");
 // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 //Set side
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
 // changeServerName();
 
}

function loveFifteen() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("0-15");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 updateGameScore();

}

function loveThirty() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("0-30");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
 updateGameScore();

}

function loveFourty() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("0-40");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 // Set side
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 // Set break point
 sheet.getRange('I' + (lastRow + 1)).setValue("1");
 updateGameScore();
 
 

}

function fifteenLove() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("15-0");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 updateGameScore();
 
}

function thirtyLove() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("30-0");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
 updateGameScore();

}

function fourtyLove() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("40-0");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 updateGameScore();

}

function fifteenAll() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("15-15");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
updateGameScore();

}

function fifteenThirty() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("15-30");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 updateGameScore();

}

function fifteenFourty() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("15-40");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
 sheet.getRange('I' + (lastRow + 1)).setValue("1");
 updateGameScore();

}

function thirtyFifteen() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("30-15");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 updateGameScore();

}

function thirtyAll() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("30-30");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
 updateGameScore();

}

function thirtyFourty() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("30-40");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 sheet.getRange('I' + (lastRow + 1)).setValue("1");
 updateGameScore();

}

function fourtyFifteen() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("40-15");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
 updateGameScore();

}

function fourtyThirty() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("40-30");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 updateGameScore();

}

function deuceAdSide() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("40-40");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Ad");
 sheet.getRange('I' + (lastRow + 1)).setValue("1");
 updateGameScore();

}

function deuceDeuceSide() {
 // Get the active sheet
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Set proper score in proper column
 lastRow = sheet.getLastRow()
 sheet.getRange('A' + (lastRow + 1)).setValue("40-40");
  // Rally length
 sheet.getRange('J' + (lastRow + 1)).setValue("1");
 // Set isPointStart to be true
 sheet.getRange('E' + (lastRow + 1)).setValue("1");
 sheet.getRange('K' + (lastRow + 1)).setValue("Deuce");
 sheet.getRange('I' + (lastRow + 1)).setValue("1");
 updateGameScore();

}

function serveZoneT() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('N' + (lastRow)).setValue('T');
}

function serveZoneWide() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
 sheet.getRange('N' + (lastRow)).setValue('Wide');
}

function serveZoneBody() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('N' + (lastRow)).setValue('Body');
}


function secondServeZoneT() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('R' + (lastRow)).setValue('T');
}

function secondServeZoneWide() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
 sheet.getRange('R' + (lastRow)).setValue('Wide');
}

function secondServeZoneBody() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('R' + (lastRow)).setValue('Body');
}

function firstServeIn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('M' + (lastRow)).setValue('1');
}

function firstServeOut() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('M' + (lastRow)).setValue('0');
}

function secondServeIn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('Q' + (lastRow)).setValue('1');
}

function secondServeOut() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnToCheck = 1; // Change this to the column number you are interested in

  // Get all values in the specified column
  var columnValues = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();

  // Iterate through the values to find the last non-empty cell
  var lastRow = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      lastRow = i + 1; // Adding 1 to convert from zero-based index to 1-based row number
    }
  }
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('Q' + (lastRow)).setValue('0');
  sheet.getRange('G' + (lastRow)).setValue('1');
  copyPointScore();
  copyGameScore();
  copyTiebreakScore();
 // whoWonPoint('error');
}


function forehandCrossDeuce() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Forehand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Crosscourt');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Deuce');


}

function forehandCrossAd() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Forehand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Crosscourt');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Ad');

}

function backhandCrossDeuce() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Backhand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Crosscourt');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Deuce');


}

function backhandCrossAd() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Backhand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Crosscourt');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Ad');

}

function forehandDTLDeuce() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Forehand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Down the Line');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Deuce');


}
function forehandDTLAd() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Forehand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Down the Line');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Ad');


}
function backhandDTLDeuce() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Backhand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Down the Line');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Deuce');
}

function backhandDTLAd() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AC' + (maxLastRow + 1)).setValue('Backhand');
  sheet.getRange('AB' + (maxLastRow + 1)).setValue('Down the Line');
  sheet.getRange('J' + (maxLastRow + 1)).setValue(sheet.getRange('J' + maxLastRow).getValue() + 1);
  sheet.getRange('K' + (maxLastRow + 1)).setValue('Ad');


}
function deuce() {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('K' + (maxLastRow)).setValue('Deuce');
}

function ad() {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('K' + (maxLastRow)).setValue('Ad');
}


function errorWideRight() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  sheet.getRange('G' + (maxLastRow)).setValue('1');
  sheet.getRange('AO' + (maxLastRow)).setValue('1');
  copyPointScore();
  copyGameScore();
  copyTiebreakScore();
//  whoWonPoint('error');
}

function errorWideLeft() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  sheet.getRange('G' + (maxLastRow)).setValue('1');
  sheet.getRange('AP' + (maxLastRow)).setValue('1');
  copyPointScore();
  copyGameScore();
  copyTiebreakScore();
 // whoWonPoint('error');
}

function errorNet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  sheet.getRange('G' + (maxLastRow)).setValue('1');
  sheet.getRange('AQ' + (maxLastRow)).setValue('1');
  copyPointScore();
  copyGameScore();
  copyTiebreakScore();
 // whoWonPoint('error');
}

function errorLong() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  sheet.getRange('G' + (maxLastRow)).setValue('1');
  sheet.getRange('AR' + (maxLastRow)).setValue('1');
  copyPointScore();
  copyGameScore();
  copyTiebreakScore();
 // whoWonPoint('error');
}

function winner() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('G' + (maxLastRow)).setValue('1');
  sheet.getRange('AN' + (maxLastRow)).setValue('1');
  copyPointScore();
  copyGameScore();
  copyTiebreakScore();
  // Check for Ace
  if ((sheet.getRange('J' + (maxLastRow)).getValue() == '1') && (sheet.getRange('AN' + (maxLastRow)).getValue() == '1')) {
      sheet.getRange('U' + (maxLastRow)).setValue('1');
  }
  //whoWonPoint("winner");
}

function slice() {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });
  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AD' + (maxLastRow)).setValue('1');
}

function volley() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AE' + (maxLastRow)).setValue('1');
}

function overhead() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AF' + (maxLastRow)).setValue('1');
}
function approach() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AG' + (maxLastRow)).setValue('1');
}

function dropshot() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AH' + (maxLastRow)).setValue('1');
}

function atNet() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var columnsToCheck = [11]; // Change this to the column number you are interested in


 var maxLastRow = 0;


 // Iterate through each specified column
 columnsToCheck.forEach(function(column) {
   // Get all values in the current column
   var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();


   // Iterate through the values to find the last non-empty cell
   for (var i = columnValues.length - 1; i >= 0; i--) {
     if (columnValues[i][0] !== "") {
       maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
       break; // Exit the loop once the last non-empty cell is found in the current column
     }
   }
 });

 // Set values in the first empty row for columns A, B, and C to zero
 sheet.getRange('AI' + (maxLastRow)).setValue('1');
}

function oppAtNet() {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var columnsToCheck = [11]; // Change this to the column number you are interested in


 var maxLastRow = 0;


 // Iterate through each specified column
 columnsToCheck.forEach(function(column) {
   // Get all values in the current column
   var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();


   // Iterate through the values to find the last non-empty cell
   for (var i = columnValues.length - 1; i >= 0; i--) {
     if (columnValues[i][0] !== "") {
       maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
       break; // Exit the loop once the last non-empty cell is found in the current column
     }
   }
 });

 // Set values in the first empty row for columns A, B, and C to zero
 sheet.getRange('AJ' + (maxLastRow)).setValue('1');
}


function lob()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in
  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('AK' + (maxLastRow)).setValue('1');

}
 
// function whoWonPoint(callerFunction) {
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//  var lastRow = sheet.getLastRow();
//  var shotInRallyValue = sheet.getRange("K" + lastRow).getValue();
//  var pointScoreValue = sheet.getRange("A" + lastRow).getValue();
//  var player1Name = sheet.getRange("BB2").getValue(); // Assuming the Player1Name column is in column AX2
//  var player2Name = sheet.getRange("BC2").getValue(); // Assuming the player1Name column is in column AY2
//  var gameScoreValue = sheet.getRange("B" + lastRow).getValue(); // Assuming gameScore column is column B
//   // Split the game score into individual scores
//  var scores = gameScoreValue.split("-");
//  var score1 = parseInt(scores[0]);
//  var score2 = parseInt(scores[1]);


// // Calculate the total games
//  var totalGames = score1 + score2;


//  var whoWonPointValue;


//  if (totalGames % 2 === 1) {
//    //switch playernames
//    var temp = player1Name;
//    player1Name = player2Name;
//    player2Name = temp;
//  }


//  // TODO: Add if else statement when gameScore = 6-6

//  if (callerFunction === "winner") {
//    whoWonPointValue = shotInRallyValue % 2 == 0 ? player2Name : player1Name;
//  } else if (callerFunction === "error") {
//    whoWonPointValue = shotInRallyValue % 2 == 0 ? player1Name : player2Name;
//  } else {
//    // Handle invalid callerFunction
//    whoWonPointValue = "Invalid callerFunction";
//  }
//   sheet.getRange("AT" + lastRow).setValue(whoWonPointValue); //put into whoWonPoint column
// }

// function changeServerName() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var lastRow = sheet.getLastRow();
//   var player1Name = sheet.getRange("BB2").getValue(); // Assuming the Player1Name column is in column AX2
//   var player2Name = sheet.getRange("BC2").getValue(); // Assuming the player1Name column is in column AY2
//   var serverName = sheet.getRange("M2").getValue();
//   var gameScoreValue = sheet.getRange("B" + (lastRow - 1)).getValue(); // get the last game score
//   if (lastRow != 2) {
//    // Split the game score into individual scores
//  var scores = gameScoreValue.split("-");
//  var score1 = parseInt(scores[0]);
//  var score2 = parseInt(scores[1]);
// // Calculate the total games
//  var totalGames = score1 + score2;
//  if (totalGames % 2 === 1) {
//  if (serverName == player1Name) {
//    sheet.getRange("M" + lastRow).setValue(player1Name); //put into serverName column
//  }
//  else {
//      sheet.getRange("M" + lastRow).setValue(player2Name); //put into serverName column
//  }  
//  } 
//  else
//  {
//  if (serverName == player1Name) {
//    sheet.getRange("M" + lastRow).setValue(player2Name); //put into serverName column
//  }
//  else {
//      sheet.getRange("M" + lastRow).setValue(player1Name); //put into serverName column
//  }  
//  }
//   }
// }

function dataError() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnsToCheck = [11]; // Change this to the column number you are interested in

  var maxLastRow = 0;

  // Iterate through each specified column
  columnsToCheck.forEach(function(column) {
    // Get all values in the current column
    var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();

    // Iterate through the values to find the last non-empty cell
    for (var i = columnValues.length - 1; i >= 0; i--) {
      if (columnValues[i][0] !== "") {
        maxLastRow = Math.max(maxLastRow, i + 1); // Update maxLastRow if a greater last row is found
        break; // Exit the loop once the last non-empty cell is found in the current column
      }
    }
  });

  // Set values in the first empty row for columns A, B, and C to zero
  sheet.getRange('BH' + (maxLastRow)).setValue('Data Error');
}
function clearAllData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Specify the row number from which you want to clear data
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - startRow + 1;

  if (numRows > 0) {
    var range = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());

    // Clear all data in the specified range
    range.clear();
  }
}




