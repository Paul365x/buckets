const ignores = [ "contents","Summary", "Totals", "Standard Transactions", "Instructions", "BucketTemplate"]; // sheet names that are not buckets
const trailerSheets = 3;  // number of non bucket sheets at the end
const startRow = 2;  // where the data begins
//const aCol = "A";    // col letter for checkboxes
const bCol = "A";    // col letter for sheet/bucket names
const cCol = "B";    // col for bucket balances
const formCol = "D"; // col for input fields range start
const fromCol = "D"; // totals column for sheet the line came from
const formDateCol = "H"; // new transaction date field
const balCol = "E";  // col in bucket for the balance
const debCol = "D";  // col in bucket for debits
const credCol = "C"; // col in buckets for credits
const folderId = ""; // id of the buckets folder
const fileId = "buckets 2024"; // id of the buckets file
const timestampLoc = "E5"; // location for the last backup timestamp
const fromCell = "D2";     // bucket selector to copy from
const toCell = "E2";       // bucket select to copy to

//https://docs.google.com/spreadsheets/d/1rXn_D5ZbQW_1B0NlzvHSeCkocHBVI-tO8MMRlpE4PUw/edit?usp=sharing
//https://docs.google.com/spreadsheets/d/1h4FB9jrM6G9b0UY0TB4Rtz4ww1Lc4RT--37hDaDzg0Y/edit?usp=sharing

//
// event handling code
//

// sets up our custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Buckets Menu' )
      .addItem('Create Sheet', 'NEWSHEET' )
      .addItem('Refresh', 'REFRESH' )
      .addItem("Apply","APPLY")
      .addItem( "Pay", "STANDARDTRANS" )
      .addItem( "Backup", "BACKUP" )
      .addItem( "Regenerate Totals", "GENERATETOTALS" )      
      .addToUi();
  resetForm();  
}

/*
// turns checkboxes into radio buttons
function onEdit(e) {
  let col = aCol.charCodeAt(0) - 64;
  if(e != null &&
    e.range.rowStart === e.range.rowEnd &&
    e.range.columnStart === col &&
    e.range.columnEnd === col) {
   //Edit occurred in column A and a single cell was edited
   if(e.oldValue === "false" && e.value === "TRUE") {
     //A checkbox was checked
     updateBucket(e.range.rowStart);
     uncheckOtherCheckboxes(e.range.rowStart);
   }
 }
}

// enforces only once checkbox ticked at a time
function uncheckOtherCheckboxes(rowToIgnore) {
 var range = SpreadsheetApp.getActive().getSheetByName("Summary").getRange(aCol + ":" + aCol );
 var maxLoc = range.getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation();
 var range = SpreadsheetApp.getActive().getSheetByName("Summary").getRange(aCol + startRow + ":" + maxLoc);
 var values = range.getValues();
 values.forEach(function (row, index) {
    if( rowToIgnore != ( index + startRow )) {    
      values[index][0] = false;
    }  
 });
 range.setValues(values);
}

// copies the selected bucket (via checkbox) to the input form
function updateBucket(row) {
  let ss = SpreadsheetApp.getActive().getSheetByName("Summary");
   var from = ss.getRange(bCol + row + ":" + bCol + row );
   var to = ss.getRange(formCol + startRow + ":" + formCol + startRow );
   let dateLoc = ss.getRange( formDateCol + startRow + ":" + formDateCol + startRow)
   to.setValue(from.getValue());
   let dateStr = splitDate( currentDate());
   dateLoc.setValue( dateStr );
}
*/
function resetForm() { 
  let ss = SpreadsheetApp.getActive().getSheetByName("Summary");
  let range = ss.getRange( formCol + startRow + ":" + formDateCol + startRow );
  range.clear(); 
  doDropdown(fromCell);
  doDropdown(toCell);  
  let dateLoc = ss.getRange( formDateCol + startRow + ":" + formDateCol + startRow)
  let dateStr = splitDate( currentDate());
  dateLoc.setValue( dateStr );
}

//
// Menu handlers
//

// Create Sheet handler
function NEWSHEET() {
    var res = SpreadsheetApp.getUi().prompt('New Sheet Name');
    var name = res.getResponseText();
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var index = activeSpreadsheet.getNumSheets() - trailerSheets;
    createSheet( name,index);    
    var summarySheet = activeSpreadsheet.getSheetByName( "Summary");
    activeSpreadsheet.setActiveSheet( summarySheet );
    REFRESH();
}

function my_notify(val) {
  var range = SpreadsheetApp.getActive().getSheetByName("Summary").getRange("I1:I1" );
  range.setValue(val);
}

// Refresh handler
function REFRESH () {
   var sheet = SpreadsheetApp.getActive().getSheetByName("Summary");
   var range = sheet.getRange(cCol+ startRow +":" + cCol );
   var maxLoc = range.getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation();
   range = sheet.getRange( bCol + startRow + ":" + maxLoc );
   range.clear();
   range.deleteCells(SpreadsheetApp.Dimension.ROWS);

   // update the names of sheets on summary page and get the last row.
   let maxRow = SETSHEETNAMES();
   if ( maxRow === 0 ) {
       return;
   }
   //SETCHECK(maxRow);
   SETBALANCE(maxRow);
}

// transaction Apply handler
function APPLY() {
  let toCol = String.fromCharCode(formCol.charCodeAt(0)+1)
  let commentCol = String.fromCharCode(formCol.charCodeAt(0) + 2);
  let amountCol = String.fromCharCode(formCol.charCodeAt(0) + 3);   
  let ss = SpreadsheetApp.getActive().getSheetByName("Summary");
  let bucketVal = ss.getRange(formCol + startRow + ":" + formCol + startRow ).getValue();
  let toVal = ss.getRange(toCol + startRow + ":" + toCol + startRow ).getValue();
  let commentVal = ss.getRange(commentCol + startRow + ":" + commentCol + startRow ).getValue();
  let amountVal = ss.getRange(amountCol + startRow + ":" + amountCol + startRow ).getValue();
  let dateStr = ss.getRange(formDateCol + startRow + ":" + formDateCol + startRow ).getValue();
  if (bucketVal !== "None") {  
    transact( bucketVal,amountVal, dateStr, commentVal );
    if (toVal !== "None") {
      amountVal *= -1;
      transact( toVal,amountVal, dateStr, commentVal );
    }
  }
  REFRESH();
  resetForm();
}

// carry out a single transaction used by apply and standardtrans
function transact( bucketVal,amountVal, dateStr, commentVal ) {
  let sheets = [];
  sheets.push( bucketVal );
  sheets.push( "Totals" );

  let credit = 0;
  let debit = 0;
  if (amountVal > 0) {
      credit = amountVal;
  } else {
      debit = amountVal;
  }    
  let values =  [ [ dateStr, commentVal, credit, debit ] ];

  sheets.forEach(function(sheetName) {
     let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
     if ( !sheet ) {
         SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
             .alert('Bucket: ' + sheetName + " doesn't exist!");
         return;
     }
     sheet.insertRowsBefore(startRow, 1);
     
     let range = sheet.getRange( bCol + startRow + ":" + debCol + startRow );
     range.setValues(values);
    
     // need to update the totals from column with sheet name
     if (sheetName == "Totals") {
        range = sheet.getRange(fromCol + startRow + ":" + fromCol + startRow )
        range.setValue( bucketVal )
     }

     range = sheet.getRange( balCol + startRow + ":" + balCol + startRow );
     range.setFormula("=" + balCol + (startRow + 1) + "+" + credCol + startRow + "+" + debCol + startRow);
     range = sheet .getRange( credCol + startRow + ":" + balCol + startRow );
     range.setNumberFormat("[Black][$$]#,##0.00;[Red][$$]-#,##0.00");  
     SORTSHEET( sheetName, fromCol );
  });
};

// make payments menu handler
function STANDARDTRANS() {
  let sheet = SpreadsheetApp.getActive().getSheetByName( "Standard Transactions" );
  let range = sheet.getRange("E1:E");
  let maxLoc = range.getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation();
  range = sheet.getRange( "A2:" + maxLoc )
  let values = range.getValues();
  let nowDate = currentDate();
  values.forEach(function (row, index) {
      let tgtDate = nextDate( row[3].toString(), row[2] );
      while( tgtDate <= nowDate ) {
          let strDate = splitDate( tgtDate );
          transact( row[4],row[1], strDate.toString(), row[0] );
          values[index][3] = tgtDate;
          tgtDate = nextDate( values[index][3].toString(), row[2] )
      }
  });
  range.setValues( values );
  REFRESH();  
}

// backup menu handler
function BACKUP() {
  let bucketFile = DriveApp.getFilesByName( fileId ).next();
  

  // get last update of buckets file
  let currentEdit = bucketFile.getLastUpdated();
  
  // make a change to the spreadsheet
  let ts = timeStamp(1);  
  let range = SpreadsheetApp.getActive().getSheetByName( "Summary" )
             .getRange( timestampLoc + ":" + timestampLoc );
  range.setValue(ts);
  
  // look for update change
  let newEdit = bucketFile.getLastUpdated();
  while ( newEdit === currentEdit ) {
      newEdit = bucketFile.getLastUpdated();
  }

  // copy file to backup
  ts.replace(/ /g,"_");
  ts = "bucket_state_" + ts;
  if (folderId != "") {
    let bucketFolder = DriveApp.getFoldersByName( folderId ).next();
  bucketFile.makeCopy(ts, bucketFolder); 
  } else {
    bucketFile.makeCopy(ts);
  }
}

// creates totals sheet by copying all transactions
function GENERATETOTALS () {
  createSheet( "Totals", 1);
  let ssTotal = SpreadsheetApp.getActive().getSheetByName( "Totals" );
  let nextRow = startRow;
  let lastRows = 1;
  let sheetNames = GETSHEETNAMES();
  sheetNames.forEach(function (sheetName) {  
    let ss = SpreadsheetApp.getActive().getSheetByName( sheetName );
    let copyRows = Number(GETMAXROW( sheetName, bCol)); 
    let from = ss.getRange( bCol + startRow + ":" + debCol + copyRows);    
    lastRows = Number(lastRows) + Number(copyRows)-1;
    if (lastRows < nextRow ) lastRows = nextRow;
    let to = ssTotal.getRange( bCol + nextRow + ":" + debCol + lastRows);
    to.setValues(from.getValues());
    let fromSheet = ssTotal.getRange(fromCol + nextRow + ":" + fromCol + lastRows);
    fromSheet.setValue(sheetName)
    nextRow = lastRows + 1;
  });

  SORTSHEET( "Totals", fromCol );
  to = ssTotal.getRange( balCol + startRow + ":" + balCol + lastRows );
  to.setFormula("=" + balCol + (startRow + 1) + "+" + credCol + startRow + "+" + debCol + startRow);

  to = ssTotal.getRange( credCol + startRow + ":" + balCol + lastRows );
  to.setNumberFormat("[Black][$$]#,##0.00;[Red][$$]-#,##0.00");
  //[Blue]#,##0;[Red]#,##0;
  //SpreadsheetApp.getActiveSheet().getRange("C2").autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
}

function blert(msg) {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert(msg);
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
}

//
// Utility functions
//

// adds a dropdown allowing selection of sheet/bucket name in the loc cell
function doDropdown( loc ) {
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( "Summary" ).getRange(loc);
  var sheets = GETSHEETNAMES();
  sheets.unshift("None");
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(sheets, true).build();
  cell.setDataValidation(rule);  
  cell.setValue("None");
}

// create a new sheet
function createSheet ( name, index ) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var yourNewSheet = activeSpreadsheet.getSheetByName( name );

  if (yourNewSheet != null) {
      activeSpreadsheet.deleteSheet(yourNewSheet);
  }
    
  var templateSh = activeSpreadsheet.getSheetByName( "BucketTemplate");    
  yourNewSheet = activeSpreadsheet.insertSheet( name, index, {template:templateSh} );
}

// returns the list of allowed sheets
function GETSHEETNAMES() {
  let ss = SpreadsheetApp.getActive();
  let sheets = ss.getSheets();
  let sheetNames =[];
  sheets.forEach(function (sheet) {
      var name = sheet.getName();
      if (!ignores.includes( name )) {
        sheetNames.push( [ sheet.getName() ]); 
      }
  });
  return sheetNames;
}

// updates the bucket list
function SETSHEETNAMES() {
  // get the list of names and populate
  let ss = SpreadsheetApp.getActive();
  let sheetNames = GETSHEETNAMES();
  
  // this is the maxRow of the list of sheet names on summary
  let sheet = ss.getSheetByName( "Summary" );
  let maxRow = startRow + sheetNames.length - 1;
  if ( maxRow < startRow) {
      return 0;
  }
  
  var range = sheet.getRange( bCol + startRow + ":" + bCol + maxRow );
  range.setValues(sheetNames);
  return maxRow;
}

/*
// update the checkbox column
function SETCHECK( maxRow ) {
   let ss = SpreadsheetApp.getActive().getSheetByName("Summary");  
   //let chkRange = ss.getActiveRange();
   ss.getRange( aCol + startRow + ":" + aCol + maxRow).insertCheckboxes();
}
*/

// sort the transactions in a bucket
function SORTSHEET( sheetName, colName ) {
   let maxRow = GETMAXROW( sheetName, bCol)
   let ss = SpreadsheetApp.getActive().getSheetByName(sheetName); 
   let totRange = ss.getRange( bCol + startRow + ":" + colName + maxRow);
   totRange.sort([{column: 1, ascending: false}]);

}

// get the last row number of a sheet and column
function GETMAXROW( sheetName, col) {
   let ss = SpreadsheetApp.getActive().getSheetByName(sheetName);  
   let range = ss.getRange( col + ":" + col );
   let maxRow = range.getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation();
   maxRow = maxRow.replace(/[^0-9]+/ig,"");
   return maxRow;
}

// update the balance column
function SETBALANCE( maxRow ) {
   var range = SpreadsheetApp.getActive().getSheetByName("Summary").getRange(bCol+ startRow +":" + bCol + maxRow);
   var values = range.getValues(); 
   values.forEach(function (row, index) {
     let crow = index + startRow;
     let balanceRange = SpreadsheetApp.getActive().getSheetByName("Summary").getRange(cCol + crow + ":" + cCol + crow );
     let bucketRange = SpreadsheetApp.getActive().
                 getSheetByName(values[index][0]).getRange(balCol + startRow + ":" + balCol + startRow );
     balanceRange.setValue( bucketRange.getValue() );
     balanceRange.setNumberFormat("$#,##0.00;$(#,##0.00)");
   });
}   
   
//
// Date manipulations
//

// return currentDate in yyyymmdd
function currentDate() {
     let d = new Date();
     return formatDate( d );
}

// convert from yyyymmdd to dd-mm-yyyy - JS native
function splitDate( date ) {
     let yr = date.substring( 0, 4 );
     let mnth = date.substring( 4, 6 );
     let day = date.substring( 6, 8 );
     return yr + "-" + mnth + "-" + day;
}

// create a date/time timestamp of now.
function timeStamp( time ) {
     let d = new Date();
     let day = d.getDate().toString();
     if ( day.length < 2 ) day = "0" + day;
     let mnth = d.getMonth() + 1;
     mnth = mnth.toString();
     if (mnth.length < 2 ) mnth = "0" + mnth;
     let yr = d.getFullYear().toString();
     let timestamp = day + "-" + mnth + "-" + yr;
     if ( time ) {
         let hrs = d.getHours().toString();
         if (hrs.length < 2) hrs = "0" + hrs;
         let mins = d.getMinutes().toString();
         if (mins.length < 2) mins = "0" + mins;
         timestamp = timestamp + " " + hrs + ":" + mins;
     }
     return timestamp;
}

// convert a date to yyyymmdd
function formatDate( date ) {
     let d = new Date( date );
     let day = d.getDate().toString();
     if ( day.length < 2 ) day = "0" + day;
     let mnth = d.getMonth() + 1;
     mnth = mnth.toString();
     if (mnth.length < 2 ) mnth = "0" + mnth;
     let yr = d.getFullYear().toString();
     return yr + mnth + day;
}

// add a week to a yyyymmdd date
function addWeekDate ( date ) {
     let dstr = splitDate( date );
     let d = new Date( dstr );
     let now = d.getDate();
     let tgt = now + 7;
     d.setDate( tgt );
     return formatDate( d );
}

// add a fortnight to a yyyymmdd date
function addFortnightDate ( date ) {
     let dstr = splitDate( date );
     let d = new Date( dstr );
     let now = d.getDate();
     let tgt = now + 14;
     d.setDate( tgt );
     return formatDate( d );
}

// add a month to a yyyymmdd date
function addMonthDate( date ) {
     let dstr = splitDate( date );
     let d = new Date( dstr );
     let now = d.getMonth();
     let tgt = now + 1;
     d.setMonth( tgt );
     return formatDate( d );
}

// add a year to a yyyymmdd date
function addYearDate( date ) {
     let dstr = splitDate( date );
     let d = new Date( dstr );
     let now = d.getFullYear();
     let tgt = now + 1;
     d.setFullYear( tgt );
     return formatDate( d );
}

function nextDate( date, type ) {
  switch(type) {
      case "M":
        return addMonthDate( date );
        break;
      case "W":
        return addWeekDate( date );
        break;
      case "F":
        return addFortnightDate( date );
        break;
      case "Y":
        return addYearDate( date );
        break;
      default:
        SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .alert('Unknown transaction period');
        return 0;
        break;
  } 
}
