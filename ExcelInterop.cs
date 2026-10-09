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
        Version version = new Version( "1.0.4" );

        Microsoft.Office.Interop.Excel.Application excelApp;
        Microsoft.Office.Interop.Excel.Workbook workbook;
        Microsoft.Office.Interop.Excel.Worksheet worksheet;

        public bool useHeader = true;
        public string fileName = "";
        public string targetFileName = "";
        public List<string[]> valuesString = new List<string[]>();
        public List<double[]> valuesDouble = new List<double[]>();
        public List<string[]> valuesTypes = new List<string[]>();
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
            }

            if ( !ok )
                Console.WriteLine("File dialog error... " );
            try
            {
                excelApp =
                    new Microsoft.Office.Interop.Excel.Application();

            }
            catch (COMException comEx)
            {
                Console.WriteLine("Excel-Interop-error: " + comEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("common error: " + ex.Message);
            }

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
                    return (false);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Allgemeiner Fehler: " + ex.Message);
                    return (false);
                }
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
                try
                {
                    workbook.SaveAs(fileName);

                }
                catch (COMException comEx)
                {
                    Console.WriteLine("Excel-Interop-Fehler: " + comEx.Message);
                    return (false);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Allgemeiner Fehler: " + ex.Message);
                    return (false);
                }
                return ( true );

            }
            return ( false );

        }   // end: DialogFileNameLoad

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


        // --------------------------------------------     the routines

        /// <summary>
        /// Reads the table's data types as array of
        /// strings into 'valuesTypes'. Needed for the analysis
        /// of foreign data.
        /// </summary>
        public void ReadTypesList( )
        {
            // die Daten auslesen
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            int startRow = usedRange.Row;
            int startCol = usedRange.Column;
            int endRow = startRow + rowCount - 1;
            int endCol = startCol + colCount - 1;

            valuesTypes = new List<string[]>();

            for ( int rowInd = startRow; rowInd < endRow; rowInd++ )
            {
                string[] temp = new string[ colCount ];
                for ( int colInd = startCol; colInd < endCol; colInd++ )
                {
                    string colName = GetColumnFromNumber( colInd );
                    temp[ colInd - startCol] =
                        worksheet.Cells[rowInd, colName].GetType().ToString()
                            ?? string.Empty;
                    valuesTypes.Add( temp );

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

        }   // end: ReadTypesList

        /// <summary>
        /// Reads the table's data as array of
        /// strings into 'valuesString'.
        /// </summary>
        public void ReadStringList( )
        {
            // die Daten auslesen
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            int startRow = usedRange.Row;
            int startCol = usedRange.Column;
            int endRow = startRow + rowCount - 1;
            int endCol = startCol + colCount - 1;

            valuesString = new List<string[]>();

            for (int rowInd = startRow; rowInd < endRow; rowInd++)
            {
                string[] temp = new string[colCount];
                for (int colInd = startCol; colInd < endCol; colInd++)
                {
                    string colName = GetColumnFromNumber(colInd);
                    temp[colInd - startCol] =
                        worksheet.Cells[rowInd, colName].ToString()
                            ?? string.Empty;
                    valuesString.Add(temp);

                }   // end: for ( int colInd

            }   // end: for ( int rowInd


        }   // end: ReadStringList

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// doubles into 'valuesDouble'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadDoubleList( )
        {
            // die Daten auslesen
            Excel.Range usedRange = excelApp.ActiveSheet.UsedRange;
            worksheet = excelApp.ActiveSheet.Worksheet;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            int startRow = usedRange.Row;
            int startCol = usedRange.Column;
            int endRow = startRow + rowCount - 1;
            int endCol = startCol + colCount - 1;

            valuesDouble = new List<double[]>();

            for (int rowInd = startRow; rowInd < endRow; rowInd++)
            {
                double[] temp = new double[colCount];
                for (int colInd = startCol; colInd < endCol; colInd++)
                {
                    string colName = GetColumnFromNumber(colInd);
                    if (worksheet.Cells[rowInd, colName].GetType() == typeof(double))
                        temp[colInd - startCol] =
                            worksheet.Cells[rowInd, colName];
                    valuesDouble.Add(temp);

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

        }   // end: ReadDoubleList

        /// <summary>
        /// Returns the number of a chosen table. A dialog will open to let you choose from
        /// the found table names.
        /// </summary>
        /// <returns>the number</returns>
        public int ReadTableNames()
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
                return( sheetNumber );

            }
            return( -1 );

        }   // end: ReadTableNames

        /// <summary>
        /// Direct query for the table name.
        /// </summary>
        /// <param name="numTable">number of the sheet</param>
        /// <returns>the name or 'string.empty'</returns>
        public string GetTableName( int numTable )
        {
            // die Daten auslesen
            Excel.Sheets sheetsList = workbook.Sheets;

            if ( (sheetsList.Count > 0)
                && (sheetsList.Count > numTable ) )
            {
                return (sheetsList[numTable].Name);  
            }
            return( string.Empty );

        }   // end: GetTableName

        /// <summary>
        /// Target file name for the writing is chosen. Produces the
        /// 'targetConnectionString' for convenience.
        /// </summary>
        /// <param name="file">already known ?</param>
        /// <param name="silent">use the dialog ?</param>
        public void ChooseTarget( ref string file, bool silent = true )
        {
            targetFileName = file;

            bool ok = false;
            if ( !silent )
                ok = DialogFileNameSave( ref targetFileName );
            if ( targetFileName != "" )
            {
                targetConnectionString =
                    conStringStart +
                    targetFileName +
                    conStringEnd +
                    withHeader;
                file = targetFileName;
            }
            else
            {
                targetConnectionString = 
                    conStringStart +
                    GetDirectory() +
                    "NewTarget.xlsx" +
                    conStringEnd +
                    withHeader;
                file = GetDirectory() + "NewTarget.xlsx";
            }
            targetFileName = file;
            //Message.Show( file );

        }   // end: ChooseTarget


        /// <summary>
        /// Intern data list double will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        public void WriteListDoubleToNewTarget( string newFileTarget = "",string newTableName = "newDoubles" )
        {
            if ( valuesDouble.Count < 1 )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return;
            }

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
            // craft the 'CREATE TABLE' and 'INSERT INTO'
            int columns =  valuesDouble[0].Length;
            string tableCreateColumns = "( ";
            string tableInsertColumns = "( ";
            switch ( columns )
            {
                case 0:
                    // no data to write
                    Message.Show( "No data to write, abort!" );
                    return;
                case 1:
                    tableCreateColumns += $"{0} DOUBLE ) ";
                    tableInsertColumns += $"{0} ) VALUES ( @0 ); ";
                    break;
                case 2:
                    tableCreateColumns += $"{0} DOUBLE, ";
                    tableCreateColumns += $"{1} DOUBLE ) ";
                    tableInsertColumns += $"{0}, {1} ) VALUES ( @0, @1 ); ";
                    break;
                default:
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableCreateColumns += $"{i} DOUBLE, ";
                    tableCreateColumns += $"{( columns - 1 )} DOUBLE );";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"{i}, ";
                    tableInsertColumns += $"{( columns - 1 )} ) VALUES ( ";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"@{i}, ";
                    tableInsertColumns += $"@{( columns - 1 )} );";
                    break;

            }
            string commandCreate = $"CREATE TABLE [{newTableName}] " 
                    + tableCreateColumns;
            Message.Show( commandCreate );
            string commandInsert = $"INSERT INTO [{newTableName}] "
                    + tableInsertColumns;
            Message.Show( commandInsert );
            using ( OleDbConnection connection = new OleDbConnection( targetConnectionString ) )
            {
                connection.Open();
                // create the table
                OleDbCommand command = new OleDbCommand( commandCreate, connection );
                command.ExecuteNonQuery();
                connection.Close();
                
                connection.Open();
                foreach ( double[] row in valuesDouble )
                {
                    command.CommandText = commandInsert;
                    command.Parameters.Clear();

                    for ( int pos = 0; pos < row.Length; pos++ )
                        //command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );
                        command.Parameters.Add( $"@{pos}", OleDbType.Double ).Value = row[ pos ];

                    command.ExecuteNonQuery();
                }

                connection.Close();
                
            }

        }   // end: WriteListDoubleToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        public void WriteListStringToNewTarget( string newFileTarget = "", string newTableName = "newStrings" )
        {
            if ( valuesString.Count < 1 )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return;
            }

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
            // craft the 'CREATE TABLE' and 'INSERT INTO'
            int columns =  valuesDouble[0].Length;
            string tableCreateColumns = "( ";
            string tableInsertColumns = "( ";
            switch ( columns )
            {
                case 0:
                    // no data to write
                    Message.Show( "No data to write, abort!" );
                    return;
                case 1:
                    tableCreateColumns += $"{0} VARCHAR ) ";
                    tableInsertColumns += $"{0} ) VALUES ( @0 ); ";
                    break;
                case 2:
                    tableCreateColumns += $"{0} VARCHAR, ";
                    tableCreateColumns += $"{1} VARCHAR ) ";
                    tableInsertColumns += $"{0}, {1} ) VALUES ( @0, @1 ); ";
                    break;
                default:
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableCreateColumns += $"{i} VARCHAR, ";
                    tableCreateColumns += $"{( columns - 1 )} VARCHAR );";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"{i}, ";
                    tableInsertColumns += $"{( columns - 1 )} ) VALUES ( ";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"@{i}, ";
                    tableInsertColumns += $"@{( columns - 1 )} );";
                    break;

            }
            string commandCreate = $"CREATE TABLE [{newTableName}] "
                    + tableCreateColumns;
            //Message.Show( commandCreate );
            string commandInsert = $"INSERT INTO [{newTableName}] "
                    + tableInsertColumns;
            //Message.Show( commandInsert );
            using ( OleDbConnection connection = new OleDbConnection( targetConnectionString ) )
            {
                connection.Open();
                // create the table
                OleDbCommand command = new OleDbCommand( commandCreate, connection );
                command.ExecuteNonQuery();
                connection.Close();

                connection.Open();
                foreach ( string[] row in valuesString )
                {
                    command.CommandText = commandInsert;
                    command.Parameters.Clear();

                    for ( int pos = 0; pos < row.Length; pos++ )
                        //command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );
                        command.Parameters.Add( $"@{pos}", OleDbType.VarChar ).Value = row[ pos ];

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }

        }   // end: WriteListStringToNewTarget

    }   // end: ExcelInterop

}   // end: namespace DbaseFrame
