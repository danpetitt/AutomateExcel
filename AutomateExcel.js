/*
  A simple JavaScript object to automate Microsoft Excel from within Windows scripts


  Usage:
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


    // Or you can do multiple files by re-using the Excel object which will be much quicker
    var excel = new Excel();
 
    excel.openFile( "d:\\testfile.xlsx" );
    excel.saveWorkSheet( 1, true );

    excel.openFile( "d:\\testfile2.xlsx" );
    excel.saveWorkSheet( 1, true );

    excel.openFile( "d:\\testfile3.xlsx" );
    excel.saveWorkSheet( 1, true );

    excel.quit();


    // You can supply a function for finding a worksheet name as the default just does a case insensitive comparison
    var excel = new Excel();
 
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
*/
 
 
var Excel = function ()
{
  var _filePath = "";
 
  var _fso = new ActiveXObject( "scripting.filesystemobject" );
 
  var _excel = new ActiveXObject( "Excel.Application" );
  var _wookbooks, _workbook;
 
  var _xlTextWindows = 20;
  var _xlUnicodeText = 42;
  var _XlSaveAsAccessModexlNoChange = 1;
  var _xlDoNotSaveChanges = 2;
 
  return {
    openFile: function ( file )
    {
      if( _excel == null )
      {
        _excel = new ActiveXObject( "Excel.Application" );
      }

      if( _file.length > 0 )
      {
        this.closeWorkbook();
      }


      _file = file;

      _excel.Visible = true;
 
      _wookbooks = _excel.Workbooks;
 
      _workbook = _wookbooks.Open( _filePath );
    },
 
    findWorkSheetNumber: function ( findFunction )
    {
      if( typeof findFunction != "function" )
      {
        var sheetName = findFunction;
        findFunction = function( worksheetName )
        {
          // Default find function does a case insensitive comparison
          return worksheetName.toLowerCase() == sheetName.toLowerCase();
        }
      }

      var sheets = _workbook.Worksheets;
 
      var total = sheets.Count;
      for ( var u = 1; u <= total ; u++ )
      {
        var ws = sheets.Item( u );
 
        ws.Activate();
 
        var worksheetName = ws.Name;

        if( findFunction( worksheetName ) )
          return u;
      }
 
      // Not found
      return -1;
    },
 
    getWorkSheet: function ( sheetNum )
    {
      var sheets = _workbook.Worksheets;
 
      var ws = sheets.Item( sheetNum );
      ws.Activate();
 
      return ws;
    },
 
    trimColumnData: function( str )
    {
      str = str.trim();
      if ( str.length > 0 && str.charAt( 0 ) == "\"" && str.charAt( str.length - 1 ) == "\"" )
      {
        str = str.substring( 1, str.length - 2 ).trim();
      }
 
      return str;
    },
 
    saveWorkSheet: function ( sheetNum, asUnicode )
    {
      if( typeof asUnicode == "undefined" )
      {
        asUnicode = true;
      }

      var sheets = _workbook.Worksheets;
 
      var ws = sheets.Item( sheetNum );
      ws.Activate();
 
      var worksheetName = ws.Name;
 
      var saveName = _filePath + "-" + sheetNum + "-(" + worksheetName + ").txt";
      saveName = saveName.replace( /\?/g, "" );
 
      // Delete old file
      if ( _fso.fileExists( saveName ) )
      {
        _fso.DeleteFile( saveName );
      }
 
      _workbook.SaveAs( saveName, asUnicode ? _xlUnicodeText : _xlTextWindows, "", "", false, false, _XlSaveAsAccessModexlNoChange );
 
      return saveName;
    },
 
    saveAllWorkSheets: function ()
    {
      var sheets = _workbook.Worksheets;
 
      var total = sheets.Count;
      var sheets = [];
      for ( var u = 1; u <= total; u++ )
      {
        sheets.push( saveWorkSheet( u ) );
      }
      return sheets;
    },
 
    closeWorkbook: function ()
    {
      if( _workbook ) _workbook.Close( false );
      _workbook = null;

      _file = "";
    },

    quit: function ()
    {
      this.closeWorkbook();
 
      if( _excel ) _excel.Quit();

      _excel = null;
    }
  }
}