AutomateExcel
=============

A simple JavaScript object to automate Microsoft Excel from within Windows scripts


###Usage
Just create a new `Excel` object, call `openFile`, `saveWorkSheet` and then `quit`.

####Methods
* `openFile( filePath );`: open the given Excel workbook file
* `saveWorkSheet( sheetNumber[, asUnicode] );`: save the specified worksheet number (1 indexed) and returns back the file path, optionally pass boolean flag to indicate whether file is saved as Unicode (UTF-16LE) Text
* `saveAllWorkSheets();`: saves all the worksheets
* `closeWorkbook();`: closes the Excel workbook, but leaves Excel running in the background
* `quit();`: closes the workbook (if one is open) and shuts down Excel
* `findWorkSheetNumber( find );`: returns a worksheet number (or -1 if unknown) that matches the `find` name passed in (case insensitive); the `find` parameter can also be a function where you can perhaps look for part of a name or make it case sensitive (look in examples below)
* `getWorkSheet( sheetNum );`: returns a Worksheet object for the given sheet number, or `null` if out of range; with this you can use some native Excel methods to manipulate the worksheet (in an example below, we delete a header row)
* `trimColumnData( str );`: helper function to trim 'quotes' from either side of the given string

####Examples
```
var excel = new Excel();

excel.openFile( "d:\\testfile.xlsx" );

// Find a named worksheet, delete the first row then save it out as UNICODE text
var sheetNum = excel.findWorkSheetNumber( "Sheet 2" );
if( sheetNum != -1 )
{
  var ws = excel.getWorkSheet( sheetNum );

  // Delete a row range
  var range = ws.Range( "A1", "AZ2" );
  var row = range.EntireRow;
  row.Delete();

  // Now save this worksheet as unicode text
  var savedTextFile = excel.saveWorkSheet( sheetNum );
}

excel.quit();
```

#####Or you can do multiple files by re-using the Excel object which will be much quicker
```
var excel = new Excel();

excel.openFile( "d:\\testfile.xlsx" );
excel.saveWorkSheet( 1, true );

excel.openFile( "d:\\testfile2.xlsx" );
excel.saveWorkSheet( 1, true );

excel.openFile( "d:\\testfile3.xlsx" );
excel.saveWorkSheet( 1, true );

excel.quit();
```

#####You can supply a function for finding a worksheet name as the default just does a case insensitive comparison
```var excel = new Excel();

excel.openFile( "d:\\testfile.xlsx" );

// Find a worksheet which contains specific text, then save it out as ANSI text
var sheetNum = excel.findWorkSheetNumber( function( workSheet ) {
  var lookFor = "My Example Sheet Number ";
  return workSheet.indexOf( lookFor.toLowerCase() ) != -1 ? true : false;
});
if( sheetNum != -1 )
{
  excel.saveWorkSheet( sheetNum, false );
}

excel.quit();
```
