/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using Excel = Microsoft.Office.Interop.Excel;

namespace DbaseFrame
{
    public class ExcelInterop
    {
        /// <summary>
        /// created on: 30.06.25
        /// last edit: 09.10.26
        /// Now since Visual Studio Pro i can use the Interop library
        /// </summary>
        Version version = new Version( "1.0.5" );

        Microsoft.Office.Interop.Excel.Application excelApp;
        Microsoft.Office.Interop.Excel.Workbook workbook;
        Microsoft.Office.Interop.Excel.Worksheet worksheet;

        public bool useHeader = true;
        public string fileName = "";
        public string targetFileName = "";
        public string[,] valuesStringArray;
        public double[,] valuesDoubleArray;
        public string[,] valuesTypesArray;
        public string[] sheets = new string[1];
        public int sheetNumber = -1;

        /// <summary>
        /// Constructor for the class. 
        /// </summary>
        /// <param name="file">a file name</param>
        /// <param name="silent">query for the name via dialog ?</param>
        public ExcelInterop( string file = "", bool silent = true,  bool doUseHeader = true )
        {
            fileName = file;
            useHeader = doUseHeader;
            bool ok = false;
            if ( !silent )
                ok = DialogFileNameLoad( ref fileName );
            try
            {
                excelApp =
                    new Microsoft.Office.Interop.Excel.Application();

            }
            catch ( COMException comEx )
            {
                Console.WriteLine( "Excel-Interop-error: " + comEx.Message );
            }
            catch ( Exception ex )
            {
                Console.WriteLine( "common error: " + ex.Message );
            }   // end: try excelApp

            // Open document
            try
            {
                // Datei öffnen
                workbook = excelApp.Workbooks.Open(
                    fileName,
                    ReadOnly: false,
                    Editable: true
                );

            }
            catch (COMException comEx)
            {
                Console.WriteLine("Excel-Interop-Fehler: " + comEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Allgemeiner Fehler: " + ex.Message);
            }   // end: try workbook

            if ( !ok )
                Console.WriteLine("File dialog error... " );

        }   // end: ExcelInterop ( constructor )

        public void ExcelInteropQuit()
        {
            // Optional: Workbook schließen
            if (workbook != null)
            {
                workbook.Close(SaveChanges: true);
                Marshal.ReleaseComObject(workbook);
            }

            // Excel beenden
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            workbook = null;
            excelApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

        }   // end: ExcelInteropQuit

        // ------------------------------ helpers

        /// <summary>
        /// Delivers the working directory with the systems separator
        /// symbol.
        /// </summary>
        /// <returns>working directory...</returns>
        string GetDirectory( )
        {
            string text =
                Directory.GetCurrentDirectory()
                + System.IO.Path.DirectorySeparatorChar;
            return ( text );

        }   // end: GetDirectory

        /// <summary>
        /// Queries a filename from the user with the standard dialog.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool DialogFileNameLoad( ref string fileName )
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = fileName; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel save file (.xlsx)|*.xlsx"; // Filter files by extension
            dialog.DefaultDirectory = GetDirectory();

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            // Process open file dialog box results
            if ( result == true )
            {
                // Open document
                fileName = dialog.FileName;
                return ( true );

            }
            return ( false );

        }   // end: DialogFileNameLoad

        /// <summary>
        /// Queries a filename from the user with the standard dialog.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool DialogFileNameSave( ref string fileName )
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.FileName = fileName; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel save file (.xlsx)|*.xlsx"; // Filter files by extension
            dialog.DefaultDirectory = GetDirectory();

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            // Process open file dialog box results
            if ( result == true )
            {
                // Save document
                fileName = dialog.FileName;
                return ( true );

            }
            return ( false );

        }   // end: DialogFileNameSave

        /// <summary>
        /// Target file name for the writing is chosen.
        /// </summary>
        /// <param name="file">already known ?</param>
        /// <param name="silent">use the dialog ?</param>
        public void ChooseTarget( ref string file, bool silent = true )
        {
            bool ok = true;
            if ( !silent )
                ok = DialogFileNameSave( ref file );
            if ( !ok )
            {
                file = GetDirectory() + "NewTarget.xlsx";
            }
            targetFileName = file;
            //Message.Show( file );

        }   // end: ChooseTarget

        /// <summary>
        /// Converts a column number to its corresponding Excel column name.
        /// </summary>
        /// <param name="col = number of the column"></param>
        /// <returns>string: Excel column name</returns>
        public string GetColumnFromNumber( int col )
        {
            string colName = "";
            while ( col > 0 )
            {
                int modulo = ( col - 1 ) % 26;
                colName = Convert.ToChar( 65 + modulo ).ToString() + colName;
                col = (int)( ( col - modulo ) / 26 );
            }
            return ( colName );

        }   // end: GetColumnFromNumber

        /// <summary>
        /// Converts a list of double arrays to a two-dimensional double array.
        /// </summary>
        /// <param name="list">The list of double arrays to convert.</param>
        /// <param name="array">The resulting two-dimensional double array.</param>
        public void ConvertListToMultiDimArray( List<double[]> list, out double[,] array )
        {
            int rows = list.Count;
            int cols = list[ 0 ].Length;
            array = new double[ cols, rows ];
            for ( int i = 0; i < cols; i++ )
            {
                for ( int j = 0; j < rows; j++ )
                {
                    array[ i, j ] = list[ j ][ i ];
                }
            }
        }   // end: ConvertListToMultiDimArray

        /// <summary>
        /// Converts a two-dimensional double array to a list of double arrays.
        /// </summary>
        /// <param name="array">The two-dimensional double array to convert.</param>
        /// <param name="list">The resulting list of double arrays.</param>
        public void ConvertMultiDimArrayToList( double[,] array, out List<double[]> list )
        {
            int rows = array.GetLength( 1 );
            int cols = array.GetLength( 0 );
            list = new List<double[]>();
            for ( int i = 0; i < rows; i++ )
            {
                double[] row = new double[ cols ];
                for ( int j = 0; j < cols; j++ )
                {
                    row[ j ] = array[ j, i ];
                }
                list.Add( row );
            }
        }   // end: ConvertMultiDimArrayToList

        // --------------------------------------------     the routines

        /// <summary>
        /// Opens a workbook with the given file name. Returns true if successful, false otherwise.
        /// </summary>
        /// <returns>boolean value indicating success or failure</returns>
        public bool OpenWorkbook()
        {
            try
            {
                // Datei öffnen
                workbook = excelApp.Workbooks.Open(
                    fileName,
                    ReadOnly: false,
                    Editable: true
                );

            }
            catch ( COMException comEx )
            {
                Console.WriteLine( "Excel-Interop-Fehler: " + comEx.Message );
                return ( false );
            }
            catch ( Exception ex )
            {
                Console.WriteLine( "Allgemeiner Fehler: " + ex.Message );
                return ( false );
            }   // end: try workbook
            return ( true );

        }   // end: OpenWorkbook

        public bool SaveWorkbook( )
        {
            try
            {
                workbook.Save();
                return ( true );
            }
            catch ( COMException comEx )
            {
                Console.WriteLine( "Excel-Interop-Fehler: " + comEx.Message );
                return ( false );
            }
            catch ( Exception ex )
            {
                Console.WriteLine( "Allgemeiner Fehler: " + ex.Message );
                return ( false );
            }   // end: try workbook

        }   // end: SaveWorkbook

        /// <summary>
        /// Saves the current workbook with a new file name. Returns true if successful, false otherwise.
        /// On success, updates the internal fileName to the new file name.
        /// </summary>
        /// <param name="newFileName"></param>
        /// <returns>boolean about the success of the operation</returns>
        public bool SaveWorkbookAs( string newFileName )
        {
            try
            {
                workbook.SaveAs( newFileName );
            }
            catch ( COMException comEx )
            {
                Console.WriteLine( "Excel-Interop-Fehler: " + comEx.Message );
                return ( false );
            }
            catch ( Exception ex )
            {
                Console.WriteLine( "Allgemeiner Fehler: " + ex.Message );
                return ( false );
            }   // end: try workbook

            fileName = newFileName;
            return ( true );

        }   // end: SaveWorkbookAs

        /// <summary>
        /// Empty the current worksheet by clearing all its cells. 
        /// This method does not delete the worksheet itself, 
        /// but removes all data and formatting from it.
        /// </summary>
        public void CleanSheet( )
        {
            worksheet = excelApp.ActiveSheet.Worksheet;
            worksheet.Cells.Clear();

        }   // end: CleanSheet

        /// <summary>
        /// Returns the number of a chosen sheet. A dialog will open to let you choose from
        /// the found sheet names.
        /// </summary>
        /// <returns>the number</returns>
        public int ReadSheetNames( )
        {
            // die Daten auslesen
            Excel.Sheets sheetsList = workbook.Sheets;

            if ( sheetsList.Count > 0 )
            {
                List<string> locSheets = new List<string>();
                sheets = new string[ sheetsList.Count ];

                foreach ( Excel.Worksheet locSheet in sheetsList )
                {
                    locSheets.Add( locSheet.Name );

                }
                sheets = locSheets.ToArray();

                DialogTablesChoice choice = new DialogTablesChoice( sheets );
                sheetNumber = choice.index;
                return ( sheetNumber );

            }
            return ( -1 );

        }   // end: ReadSheetNames

        /// <summary>
        /// Direct query for the sheet name.
        /// </summary>
        /// <param name="numTable">number of the sheet</param>
        /// <returns>the name or 'string.empty'</returns>
        public string GetSheetName( int numTable )
        {
            // die Daten auslesen
            Excel.Sheets sheetsList = workbook.Sheets;

            if ( ( sheetsList.Count > 0 )
                && ( sheetsList.Count > numTable ) )
            {
                return ( sheetsList[ numTable ].Name );
            }
            return ( string.Empty );

        }   // end: GetSheetName

        /// <summary>
        /// Sets the name of a chosen sheet.
        /// </summary>
        /// <param name="numTable"></param>
        /// <param name="newName"></param>
        public void SetSheetName( int numTable, string newName )
        {
            // die Daten auslesen
            Excel.Sheets sheetsList = workbook.Sheets;
            if ( ( sheetsList.Count > 0 )
                && ( sheetsList.Count > numTable ) )
            {
                sheetsList[ numTable ].Name = newName;
            }

        }   // end: SetSheetName

        /// <summary>
        /// Creates a new sheet in the workbook.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns>the number of the new sheet</returns>
        public int NewSheet( string sheetName = "" )
        {
            Excel.Worksheet newSheet = workbook.Sheets.Add();
            newSheet.Name = sheetName;

            return ( newSheet.Index );

        }   // end: NewSheet

        /// <summary>
        /// Reads the table's data types as 2d-array of
        /// strings into 'valuesTypesArray'. Needed for the analysis
        /// of foreign data.
        /// </summary>
        public void ReadTypes( )
        {
            // get the dimensions from Excel and create the array
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            int startRow = 1;
            int startCol = 1;
            int endRow = startRow + rowCount - 1;
            int endCol = startCol + colCount - 1;

            valuesTypesArray = new string[ colCount, rowCount ];

            string xDim1 = GetColumnFromNumber( startCol );
            string xDim2 = GetColumnFromNumber( endCol );
            string yDim1 = startRow.ToString();
            string yDim2 = rowCount.ToString();

            // a dream would be the next line, but it is not save for input
            // valuesTypesArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as string[,];

            for ( int rowInd = startRow; rowInd < endRow; rowInd++ )
            {
                string temp;
                for ( int colInd = startCol; colInd < endCol; colInd++ )
                {
                    string colName = GetColumnFromNumber( colInd );
                    // beware: the cell index works like this coming line
                    temp =
                        worksheet.Cells[ rowInd, colName].GetType().ToString()
                            ?? string.Empty;

                    valuesTypesArray[ colInd, rowInd ] =  temp ;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

        }   // end: ReadTypes

        /// <summary>
        /// Reads the table's data as 2d-array of
        /// strings into 'valuesStringArray'.
        /// </summary>
        public void ReadStrings()
        {
            // get the dimensions from Excel and create the array
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            int startRow = 1;
            int startCol = 1;
            int endRow = startRow + rowCount - 1;
            int endCol = startCol + colCount - 1;

            valuesStringArray = new string[ colCount, rowCount ];

            // a dream would be the next line, but it is not save for input
            // valuesStringArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as string[,];

            for ( int rowInd = startRow; rowInd <= endRow; rowInd++)
            {
                string temp;
                for (int colInd = startCol; colInd <= endCol; colInd++)
                {
                    string colName = GetColumnFromNumber(colInd);
                    // beware: the cell index works like this coming line
                    temp =
                        worksheet.Cells[rowInd, colName].ToString()
                            ?? string.Empty;

                    valuesStringArray[ colInd, rowInd ] = temp;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

        }   // end: ReadStrings

        /// <summary>
        /// Writes the 2d-array to the active sheet.
        /// Overwrites anything in its way and starts at A1.
        /// </summary>
        public void WriteStrings()
        {
            // calculate the Excel range to write to
            int colCount = valuesStringArray.GetLength( 0 );
            int rowCount = valuesStringArray.GetLength( 1 );

            int startCol = 1;
            int endCol = startCol + colCount - 1;
            int startRow = 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startRow );
            string xDim2 = GetColumnFromNumber( endRow );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = valuesStringArray;

        } // end: WriteStrings

        /// <summary>
        /// Reads a subset of the table's data as a 2D array of strings.
        /// </summary>
        /// <param name="startX">upper left X coordinate</param>
        /// <param name="startY">upper left Y coordinate</param>
        /// <param name="endX">lower right X coordinate</param>
        /// <param name="endY">lower right Y coordinate</param>
        /// <returns>subset of the table's data as a 2D array of strings</returns>
        public string[,] ReadStrings( int startX, int startY, int endX, int endY )
        {
            // get the dimensions from Excel and create the array
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int startRow = startX;
            int startCol = startY;
            int endRow = endX;
            int endCol = endY;

            string[,] locStringArray = new string[ endX - startX + 1, endY - startY + 1 ];

            // a dream would be the next line, but it is not save for input
            // valuesStringArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as string[,];

            for ( int rowInd = startRow; rowInd <= endRow; rowInd++ )
            {
                string temp;
                for ( int colInd = startCol; colInd <= endCol; colInd++ )
                {
                    string colName = GetColumnFromNumber(colInd);
                    // beware: the cell index works like this coming line
                    temp =
                        worksheet.Cells[ rowInd, colName ].ToString()
                            ?? string.Empty;

                    locStringArray[ colInd - startX, rowInd - startY ] = temp;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

            return( locStringArray );

        }   // end: ReadStrings

        /// <summary>
        /// Writes the 2d-array to the active sheet at the specified coordinates.
        /// Overwrites anything in its way.
        /// </summary>
        public void WriteStrings( string[,] dataArray, int startX, int startY )
        {
            // calculate the Excel range to write to
            int colCount = dataArray.GetLength( 0 );
            int rowCount = dataArray.GetLength( 1 );

            int startCol = startX;
            int endCol = startCol + colCount - 1;
            int startRow = startY;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startRow );
            string xDim2 = GetColumnFromNumber( endRow );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = dataArray;

        } // end: WriteStrings

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// doubles into 'valuesDoubleArray'.
        /// </summary>
        public void ReadDoubles()
        {
            // get the dimensions from Excel and create the array
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            int startRow = 1;
            int startCol = 1;
            int endRow = startRow + rowCount - 1;
            int endCol = startCol + colCount - 1;

            valuesDoubleArray = new double[ colCount, rowCount ];

            // a dream would be the next line, but it is not save for input
            // valuesDoubleArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as double[,];

            for (int rowInd = startRow; rowInd <= endRow; rowInd++)
            {
                double temp = 0;
                for (int colInd = startCol; colInd <= endCol; colInd++)
                {
                    string colName = GetColumnFromNumber(colInd);
                    // beware: the cell index works like this coming lines
                    if ( worksheet.Cells[rowInd, colName].GetType() == typeof(double))
                        temp =
                            worksheet.Cells[rowInd, colName];

                    valuesDoubleArray[ colInd, rowInd ] = temp;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

        }   // end: ReadDoubles

        /// <summary>
        /// Writes the 2d-array to the active sheet.
        /// Overwrites anything in its way and starts at A1.
        /// </summary>
        public void WriteDoubles( )
        {
            // calculate the Excel range to write to
            int colCount = valuesDoubleArray.GetLength( 0 );
            int rowCount = valuesDoubleArray.GetLength( 1 );

            int startCol = 1;
            int endCol = startCol + colCount - 1;
            int startRow = 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startRow );
            string xDim2 = GetColumnFromNumber( endRow );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = valuesDoubleArray;

        } // end: WriteDoubles

        /// <summary>
        /// Reads a subset of the table's data as a 2D array of doubles.
        /// </summary>
        /// <param name="startX">upper left X coordinate</param>
        /// <param name="startY">upper left Y coordinate</param>
        /// <param name="endX">lower right X coordinate</param>
        /// <param name="endY">lower right Y coordinate</param>
        /// <returns>subset of the table's data as a 2D array of doubles</returns>
        public double[,] ReadDoubles( int startX, int startY, int endX, int endY )
        {
            // get the dimensions from Excel and create the array
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int startRow = startX;
            int startCol = startY;
            int endRow = endX;
            int endCol = endY;

            double[,] locDoubleArray = new double[ endX - startX + 1, endY - startY + 1 ];

            // a dream would be the next line, but it is not save for input
            // valuesDoubleArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as double[,];

            for ( int rowInd = startRow; rowInd <= endRow; rowInd++ )
            {
                double temp = 0;
                for ( int colInd = startCol; colInd <= endCol; colInd++ )
                {
                    string colName = GetColumnFromNumber(colInd);
                    // beware: the cell index works like this coming line
                    temp =
                        worksheet.Cells[ rowInd, colName ]
                            ?? 0.0;

                    locDoubleArray[ colInd - startX, rowInd - startY ] = temp;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

            return ( locDoubleArray );

        }   // end: ReadDoubles

        /// <summary>
        /// Writes the 2d-array to the active sheet at the specified coordinates.
        /// Overwrites anything in its way.
        /// </summary>
        public void WriteDoubles( double[,] dataArray, int startX, int startY )
        {
            // calculate the Excel range to write to
            int colCount = dataArray.GetLength( 0 );
            int rowCount = dataArray.GetLength( 1 );

            int startCol = startX;
            int endCol = startCol + colCount - 1;
            int startRow = startY;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startRow );
            string xDim2 = GetColumnFromNumber( endRow );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = dataArray;

        } // end: WriteDoubles

        /// <summary>
        /// Intern 2d-array of doubles will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        public void WriteDoublesToNewTarget( string newFileTarget = "",string newTableName = "newDoubles" )
        {
            bool overwrite = false;
            while ( !overwrite )
            {
                if ( newFileTarget == "" )
                    ChooseTarget( ref newFileTarget, false );
                else
                    ChooseTarget( ref newFileTarget, true );

                if ( File.Exists( newFileTarget ) )
                {
                    overwrite = Message.Ask( "Do you want to delete the file and its contents ?" );
                    if ( overwrite )
                        File.Delete( newFileTarget );

                }
                else
                    overwrite = true;
                if ( !overwrite )
                    newFileTarget = "";

            }
            Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets[ 1 ].Name = newTableName;


            // calculate the Excel range to write to
            int colCount = valuesDoubleArray.GetLength( 0 );
            int rowCount = valuesDoubleArray.GetLength( 1 );

            int startCol = 1;
            int endCol = startCol + colCount - 1;
            int startRow = 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startRow );
            string xDim2 = GetColumnFromNumber( endRow );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = valuesDoubleArray;

            SaveWorkbook();

        }   // end: WriteDoublesToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        public void WriteStringsToNewTarget( string newFileTarget = "", string newTableName = "newStrings" )
        {
            bool overwrite = false;
            while ( !overwrite )
            {
                if ( newFileTarget == "" )
                    ChooseTarget( ref newFileTarget, false );
                else
                    ChooseTarget( ref newFileTarget, true );

                if ( File.Exists( newFileTarget ) )
                {
                    overwrite = Message.Ask( "Do you want to delete the file and its contents ?" );
                    if ( overwrite )
                        File.Delete( newFileTarget );

                }
                else
                    overwrite = true;
                if ( !overwrite )
                    newFileTarget = "";

            }
            Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets[ 1 ].Name = newTableName;


            // calculate the Excel range to write to
            int colCount = valuesStringArray.GetLength( 0 );
            int rowCount = valuesStringArray.GetLength( 1 );

            int startCol = 1;
            int endCol = startCol + colCount - 1;
            int startRow = 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startRow );
            string xDim2 = GetColumnFromNumber( endRow );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = valuesStringArray;

            SaveWorkbook();

        }   // end: WriteStringsToNewTarget

    }   // end: ExcelInterop

}   // end: namespace DbaseFrame
