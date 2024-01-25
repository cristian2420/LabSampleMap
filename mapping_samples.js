// Main function to map clinical data to tank sheet
function conditionalMapping(clinicalSheetName = "Cancer", tankSheetName = "TANK 1", tankNumber = 1, verbose = true, color = "green") {
  // Get data from spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tankSheetName);
  var clinicalSheet = ss.getSheetByName(clinicalSheetName); 
  var clinicalSheetIdx = clinicalSheet.getSheetId();   
  // Filter data in tank1
  var colTank = getColumnIndexByName(clinicalSheet, "Tank");

  if (colTank === -1) {
    Logger.log("Column not found: Tank");
    return;
  }

  // Index of columns [Rack, Box, Position]
  var posCols = ["Rack", "Box", "Position"];
  var idxCols = posCols.map(columnName => getColumnIndexByName(clinicalSheet, columnName));
  // Column to check if samples has been used
  var usedCol = getColumnIndexByName(clinicalSheet, "Used?");
  if (usedCol === -1) {
    Logger.log("Column not found: Used?");
    return;
  }

  // Batch read data from clinical sheet
  var clinicalDataRange = clinicalSheet.getRange(2, 1, clinicalSheet.getLastRow() - 1, clinicalSheet.getLastColumn());
  var clinicalData = clinicalDataRange.getValues();
  // Column values from clinicaldata
  var rowsTank = clinicalData.map(function(row) {
    return row[colTank - 1];
  });
  
  // Batch read data from tank sheet
  var tankDataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var tankData = tankDataRange.getValues();
  // Filter rows with column "Tank" == 1 and paste them with their values in columns [Rack, Box, Position]
  var coordinates = {};

  for (var i = 1; i <= clinicalData.length; i++) {
    if (rowsTank[i] === tankNumber) {
      var rowData = [clinicalData[i][idxCols[0] - 1], clinicalData[i][idxCols[1] - 1], clinicalData[i][idxCols[2] - 1]];
      // Check if all values are valid numbers
      var validArray = rowData.every(function (element) {
        return element !== "" && typeof element === 'number' && isFinite(element);
      });

      // Checked if value in column "Used?" == "No"
      var usedsample = clinicalData[i][usedCol - 1];
      if (!validArray || usedsample !== "No" ) {
        continue;
      };
      // Check if rack is already a key in the dictonary
      if (!coordinates.hasOwnProperty(rowData[0])){
        coordinates[rowData[0]] = {}
      };
      // Check if Box is already a key in the nested dictonary
      if (!coordinates[rowData[0]].hasOwnProperty(rowData[1])){
        coordinates[rowData[0]][rowData[1]] = {};
      };
      // Check if position is already a value in the nested dictonary
      if (!coordinates[rowData[0]][rowData[1]].hasOwnProperty(rowData[2])){
        coordinates[rowData[0]][rowData[1]][rowData[2]] = {};
      };
      // Add the A1 notation as value to map from source
      coordinates[rowData[0]][rowData[1]][rowData[2]] = A1notation(i+1, idxCols[2]-1)
    };
  }
  
  // Racks are in the first row. Get indeces of unique racks
  var firstRowValues = tankData[0];

  // Iterate by rack keys
  var orderedRacks = Object.keys(coordinates).sort(function(a,b){
    return parseInt(a) - parseInt(b);
  });
  for (var i = 0; i < orderedRacks.length; i++) {
    var rackValue = orderedRacks[i];
    var rackIdx = firstRowValues.indexOf("Rack " + rackValue) + 1;
    if (verbose){
      Logger.log("Looking at Rack: " + rackValue);
      Logger.log("RackIdx: " + rackIdx);
    };
    // Check if rack is found
    if (rackIdx === 0) {
      Logger.log("Rack not found: " + rackValue);
      continue;
    };
    
    // Look for boxes in the same rack
    var orderedBoxes = Object.keys(coordinates[rackValue]).sort(function(a,b){
      return parseInt(a) - parseInt(b);
    });
    var rackColValues = tankData.map(function (row) {
      return row[rackIdx - 1];
    })
    
    for (var j = 0; j < orderedBoxes.length; j++){
      var boxValue = orderedBoxes[j];
      var boxIdx = rackColValues.indexOf("Box " + boxValue) + 1;
      if (verbose){
        Logger.log("Looking at Box: " + boxValue);
        Logger.log("BoxIdx: " + boxIdx);
      };
      if (boxIdx === 0){
        Logger.log("Box not found");
        continue;
      };
      // Looking at positions in the same box
      var orderedPositions = Object.keys(coordinates[rackValue][boxValue]).sort(function(a, b){
        return parseInt(a) - parseInt(b);
      });
      for (var k = 0; k < orderedPositions.length; k++){
        var posValue = orderedPositions[k];
        if (verbose){
          Logger.log("Looking at Position: " + posValue);
        };
        var rowPos = Math.ceil(Number(posValue)/10); // Get the row index of the position
        var colPos = Number(posValue)%10; // Get the column index of the position
        if (colPos === 0) {
          colPos = 10;
        };

        rowPos = rowPos + boxIdx;
        colPos = (rackIdx - 1) + (colPos);
        if (verbose){
          Logger.log("Looking at Position: " + posValue + "\nCell to color: (" + rowPos + "," + colPos + ")");
        };
        var clinicalLinkFormula = '=HYPERLINK("#gid=' + clinicalSheetIdx +
                                  '&range=' + coordinates[rackValue][boxValue][posValue] + '","' + posValue + '")';
        // Set the hyperlink formula to the cell and change color
        sheet.getRange(rowPos, colPos).setFormula(clinicalLinkFormula).setBackground(color);
      }
    }
  }
}

// Helper function to get column indices by name
function getColumnIndexByName(sheet, columnName) {
  var headerRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headerRowValues.length; i++) {
    if (headerRowValues[i].trim() == columnName) {
      return i + 1;
    }
  }
  Logger.log("Column not found:" + columnName);
  return -1; // Return a default value (e.g., -1) to indicate that the column was not found
}

// Helper function to get A1 notation from row and column index
// Adapt from Amit Agarwal https://www.labnol.org/convert-column-a1-notation-210601 
function A1notation(row = 3852, col = 19) {
  var a1Notation = [row + 1];
  var totalAlphabets = 'Z'.charCodeAt() - 'A'.charCodeAt() + 1;
  var block = col;
  while (block >= 0) {
    a1Notation.unshift(String.fromCharCode((block % totalAlphabets) + 'A'.charCodeAt()));
    block = Math.floor(block / totalAlphabets) - 1;
  }
  return a1Notation.join('');
}

function resetCellColors(sheet) {
  // Get the dimensions of the sheet
  var numRows = sheet.getLastRow();
  var numCols = sheet.getLastColumn();

  // Reset cell colors for each cell in the sheet
  var columnRack = 0;
  for (var row = 1; row <= numRows; row++) {
    for (var col = 1; col <= numCols; col++) {
      // if we are in a cell with Rack or box in it, skip two
      var cellValue = sheet.getRange(row, col).getValue();
      if (cellValue === "Rack") {
        // Skip two rows
        row += 2;
        columnRack = 0;
      } else if (cellValue === "Box" || cellValue === null || cellValue.trim() === "") {
        // Skip one row
        row += 1;
      }
      //
      if (columnRack === 10){
        col += 1;
      }
      sheet.getRange(row, col).setBackground(null);
      columnRack += 1;
    }

  }
}

function clinicalDataset() {
  // Array of sample sheet names
  var sheetToTankMapping = {
    "Cancer": { tanks: ["TANK 1"], color: "#EAC7C7" },
    "Asthma": { tanks: ["TANK 1"], color: "#D5E3E8" },
    "CCHI": { tanks: ["TANK 1"], color: "E8A2A2" },
    "HIPC": { tanks: ["TANK 1","TANK 2"], color: "#F7F5EB" },
    "LJI PBMC": { tanks: ["TANK 2"], color: "#A0C3D2" },
    "DICE LCL": { tanks: ["TANK 1"], color: "#FAEDCB" },
    "DICE PBMC STOCK": { tanks: ["TANK 1","TANK 2"], color: "#C9E4DE" },
    "DICE PBMC BACKUP": { tanks: ["TANK 2"], color: "#C6DEF1" },
    "DICE patients": { tanks: ["TANK 2"], color: "#DBCDF0" },
    "IM-TCR": { tanks: ["TANK 1"], color: "#F2C6DE" },
    "Personal": { tanks: ["TANK 1"], color: "#F7D9C4" },
    // Add more mappings as needed
  };
  return sheetToTankMapping;
}

function runForMultipleSheets() {
  // Array of sample sheet names
  var sheetToTankMapping = clinicalDataset();
  
  // Loop through each sample sheet name
  for (var sheetName in sheetToTankMapping) {
    var tankInfo = sheetToTankMapping[sheetName];
    var tankNames = tankInfo.tanks;
    var tankColor = tankInfo.color;
    // Execute the conditionalMapping_V2 function for the current sample sheet and tanks
    for (var i = 0; i < tankNames.length; i++) {
      var tankName = tankNames[i];
      var tankNumber = Number(tankName.replace("TANK ", ""));
      Logger.log(sheetName + " " + tankName + " " + " " + tankNumber);
      conditionalMapping(sheetName, tankName, tankNumber, false, tankColor);
    }
  }
}


// Trigger update on edit
// ON CONSTRUCTION
// DO NOT USE IT YET
function onEdit(e) { 
  const sheetToTankMapping = clinicalDataset();
  const clinicalSheets = Object.keys(sheetToTankMapping)
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const as = e.source.getActiveSheet();
  const oldvalue = e.oldvalue();
  const newValue = e.value();

  // Check if edit happen in just one cell
  if (length(row) > 1){
    return;
  }
  // Check if edit happen in any of the clinical sheets
  if (!clinicalSheets.includes(as.getName())){
    return;
  }

  // Check if edit happen in any of the below columns
  const impCols = ["Tank", "Rack", "Box", "Position", "Used?"];
  const clinicalData = as.getRange(1, 1, as.getLastRow(), as.getLastColumn());
  const header = clinicalData[0];
  const impIdx = impCols.map(function(element) {
    return header.indexOf(element);
  });

  if (!impIdx.includes(col - 1)){
    return;
  };
  // if edit was on any of those columns check for their values
  const modifiedRow = clinicalData[row - 1];
  const TankValue = modifiedRow[impIdx[0]];
  const rackValue = modifiedRow[impIdx[0]];
  const boxValue = modifiedRow[impIdx[0]];
  const positionValue = modifiedRow[impIdx[0]];
  const usedValue = modifiedRow[impIdx[0]];

  // If used? was modify from No to Yes reset the cell in the tank sheet
  if ((col - 1) == impIdx[4]){
    if (oldvalue == "Yes" && newValue == "No"){
      
    } else if(oldvalue == "No" && newValue == "Yes"){
      
    } else {
      Logger.Log();
    }
  } else {

  }


}
