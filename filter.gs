function onOpen(){
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Filter region', functionName: 'testRegion'},
    {name: 'Filter job role', functionName: 'testJobRole'},
    {name: 'Filter team', functionName: 'testTeam'},
    {name: 'Full Filter (EMEA, LCS, Employee)', functionName: 'fullTest'}
  ];
  spreadsheet.addMenu("Filters", menuItems)
}

function filterColumn(sheet, columnHeader, targetValue){
  /* 
  sheet: The sheet object which needs to be filtered.
  columnHeader: The header of the column which is being filtered
  targetValue: The value which should be kept in the trix - the current person's role/region/etc.
  
  Given a sheet, column header and value, filters a column to remove irrelevant rows based
  on the rules in the isMatch() function.
  */

  var headerCell
  
  for(var i = 1; i < sheet.getMaxRows(); i++){
    var cell = sheet.getRange(2, i, 1, 1)
    var cellContents = cell.getValue().toLowerCase().trim()
    var trimmedHeader = columnHeader.toLowerCase().trim()
    Logger.log("Cell Contents: " + cellContents + ".\nHeader: "+ trimmedHeader +".")
    if(cellContents === trimmedHeader){
      headerCell = cell
      break
    }
  }
  
  var headerCellRow = headerCell.getRow()
  var columnIndex = headerCell.getColumn()
    
  var i = headerCellRow + 1
    
  while(!sheet.getRange(i, columnIndex).isBlank()){
    var currentCell = sheet.getRange(i, columnIndex)
    if(!isMatch(currentCell.getValue(), targetValue)){
      sheet.deleteRow(i)    
    }
    else i++
  }
}
    
function isMatch(input, targetValue){
  /*
  input: Value from cell in trix
  targetValue: Current property to be matched & kept in trix
  
  Takes 2 values, sets both to lowercase and trims. If the input contains the value, or the input is 
  "all", return true. else return false.
  */
  
  var trimmedInput = input.toLowerCase().trim()
  var trimmedValue = targetValue.toLowerCase().trim()
  if(trimmedInput === "all"){
    return true
  }
  else{
    var split = trimmedInput.split("|")
    .map(function(word) {
      word = word.trim()
      return word 
    })
    if(split.indexOf(trimmedValue) !== -1){
      return true
    } 
    return false
  }
}
    
function testIsMatch(){
  Logger.log("Expected output: True. Actual output: " + isMatch("all", "EMEA"))
  Logger.log("Expected output: True. Actual output: " + isMatch("EMEA", "EMEA"))
  Logger.log("Expected output: False. Actual output: " + isMatch("APAC", "EMEA"))    
  Logger.log("Expected output: True. Actual output: " + isMatch("EMEA | APAC", "EMEA"))
  Logger.log("Expected output: True. Actual output: " + isMatch("   EMEA   |   APAC   ", " APAC  "))
}

function testRegion(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1").copyTo(SpreadsheetApp.getActive()).setName("Sheet2")
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2")
  filterColumn(newSheet, "Region", "EMEA")
}

function testJobRole(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1").copyTo(SpreadsheetApp.getActive()).setName("Sheet3")
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet3")
  filterColumn(newSheet, "Job Role", "Intern")
}

function testTeam(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1").copyTo(SpreadsheetApp.getActive()).setName("Sheet4")
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet4")
  filterColumn(newSheet, "Team", "Platform")
}

function fullTest(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1").copyTo(SpreadsheetApp.getActive()).setName("Sheet5")
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet5")
  filterColumn(newSheet, "Region", "EMEA")
  filterColumn(newSheet, "Job Role", "Employee")
  filterColumn(newSheet, "Team", "LCS")
}
