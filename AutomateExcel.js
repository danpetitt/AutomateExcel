/**
 * AutomateExcel - A simple JavaScript object to automate Microsoft Excel from within Windows scripts
 * @author Dan Petitt <danp@coderanger.com>
 * Homepage: https://github.com/coderangerdan/AutomateExcel
 *
 * The MIT License (MIT)
 * Copyright (c) 2014 Dan Petitt
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
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