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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Controls;
using System.IO;
using System.Net;
using System.Windows.Media.Animation;
using System.Configuration;
using System.Data;

namespace DbaseFrame
{
    public class ExcelOleDb
    {
        /// <summary>
        /// created on: 22.10.24
        /// last edit: 27.11.26
        /// 27.11.26: Lesen über OpenSource( ): fehlende Datei → FileNotFoundException ( vorher legte der
        ///           ACE-Treiber stillschweigend eine neue, leere Mappe an ), defekte Datei →
        ///           InvalidDataException mit verständlicher Meldung statt roher OleDbException;
        ///           relativer Dateiname → voller Pfad; Lesen ohne gewählte Tabelle → klare Meldung;
        ///           WriteListStringToNewTarget zählt die Spalten der Texte ( vorher der Zahlen )
        /// </summary>
        Version version = new Version( "1.0.14" );

        // Connect to the Excel file
        public string conStringStart =
            "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=";
        public string conStringEnd =
            ";Extended Properties=\"Excel 12.0 Xml;";
        public string withHeader = "HDR=YES;\"";
        public string withoutHeader = "HDR=NO;\"";
        public bool useHeader = true;
        public string connectionString = "";
        public string targetConnectionString = "";
        public string fileName = "";
        public string targetFileName = "";
        public List<string[]> valuesString = new List<string[]>();
        public List<double[]> valuesDouble = new List<double[]>();
        public List<string[]> valuesTypes = new List<string[]>();
        public string[] sheets = new string[1];
        public int sheetNumber = -1;

        /// <summary>
        /// Datei, aus der gelesen wird ( fileName oder ohne Namen die Demo-Datei Parable_Demo.xlsx )
        /// </summary>
        string sourceFile = "";

        /// <summary>
        /// Constructor for the class.
        /// </summary>
        /// <param name="file">a file name</param>
        /// <param name="silent">query for the name via dialog ?</param>
        public ExcelOleDb( string file = "", bool silent = true,  bool doUseHeader = true )
        {
            fileName = file;
            useHeader = doUseHeader;
            bool ok = false;
            if ( !silent )
                ok = DialogFileNameLoad( ref fileName );
            if ( fileName != "" )
            {
                sourceFile = Path.GetFullPath( fileName );  // relativer Name → voller Pfad, eindeutig
                connectionString =
                    conStringStart + sourceFile + conStringEnd;
                if ( useHeader )
                    connectionString += withHeader;
                else 
                    connectionString += withoutHeader;
            }
            else
            {
                sourceFile = GetDirectory() + "Parable_Demo.xlsx";
                connectionString =
                    conStringStart +
                    GetDirectory() +
                    "Parable_Demo.xlsx" +
                    conStringEnd +
                    withoutHeader;

            }

        }   // end: ExcelOleDb ( constructor )

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
                // Open document
                fileName = dialog.FileName;
                return ( true );

            }
            return ( false );

        }   // end: DialogFileNameLoad

        /// <summary>
        /// Öffnet die Verbindung zur Quelldatei zum Lesen.
        /// </summary>
        /* Die Datei hat zu passen, sonst Fehlermeldung:
         *  - fehlt sie, legt der ACE-Treiber beim Öffnen stillschweigend eine neue, leere Mappe an
         *    → vorher prüfen, FileNotFoundException
         *  - ist sie beschädigt, kein Excel-Format, kennwortgeschützt oder woanders gesperrt, wirft Open( ) eine OleDbException
         *    ( "Die externe Tabelle hat nicht das erwartete Format." ) → InvalidDataException mit Dateinamen
        */
        /// <returns>die geöffnete Verbindung ( mit using verwenden )</returns>
        OleDbConnection OpenSource( )
        {
            if ( !File.Exists( sourceFile ) )
                throw new FileNotFoundException( $"Excel-Datei nicht gefunden: {sourceFile}", sourceFile );
            OleDbConnection conn = new OleDbConnection( connectionString );
            try
            {
                conn.Open( );
            }
            catch ( OleDbException ex )
            {   // beschädigt oder kein Excel-Format
                conn.Dispose( );
                throw new InvalidDataException(
                    $"Excel-Datei kann nicht gelesen werden ( beschädigt, kein Excel-Format, kennwortgeschützt oder woanders geöffnet ): {sourceFile}\n"
                    + ex.Message, ex );
            }   // Ende: beschädigt oder kein Excel-Format
            return ( conn );

        }   // Ende: OleDbConnection OpenSource

        /// <summary>
        /// Name der gewählten Tabelle für das Lesen. Ohne gültige Wahl eine klare Fehlermeldung
        /// ( vorher IndexOutOfRangeException ).
        /// </summary>
        /// <returns>Tabellenname, z. B. "Tabelle1$"</returns>
        string SelectedTable( )
        {
            if ( ( sheetNumber < 0 ) || ( sheetNumber >= sheets.Length ) || string.IsNullOrEmpty( sheets[ sheetNumber ] ) )
                throw new InvalidOperationException(
                    "Keine Tabelle gewählt: zuerst ReadTableNames( ) aufrufen oder 'sheets' und 'sheetNumber' setzen." );
            return ( sheets[ sheetNumber ] );

        }   // Ende: string SelectedTable

        // --------------------------------------------     the routines

        /// <summary>
        /// Reads the table's data types as anonymous array of
        /// strings into 'valuesTypes'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadTypesList( )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                OleDbCommand  command = new OleDbCommand ( $"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                
                valuesTypes = new List<string[]>();

                while ( reader.Read() )
                {
                    int cols = reader.FieldCount;
                    string[] temp = new string[ cols ];
                    for ( int pos = 0; pos < cols; pos++ )
                        temp[ pos ] =
                            reader[ pos ].GetType().ToString()
                            ?? string.Empty;
                    valuesTypes.Add( temp );

                }
                reader.Close();
                conn.Close();

            }   // end: using

        }   // end: ReadTypesList

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// strings into 'valuesString'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadStringList( )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                OleDbCommand command = new OleDbCommand( $"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                
                valuesString = new List<string[]>();

                while ( reader.Read() )
                {
                    string[] temp = new string[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        temp[ pos ] = 
                            reader[ pos ].ToString()
                            ?? string.Empty;
                    valuesString.Add( temp );

                }
                
            }   // end: using

        }   // end: ReadStringList

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// doubles into 'valuesDouble'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadDoubleList( )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                OleDbCommand command = new OleDbCommand($"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                
                valuesDouble = new List<double[]>();

                while ( reader.Read() )
                {
                    double[] temp = new double[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        if ( reader[ pos ].GetType() == typeof( double ) )
                            temp[ pos ] = reader.GetDouble( pos );
                    valuesDouble.Add( temp );

                }

            }   // end: using

        }   // end: ReadDoubleList

        /// <summary>
        /// Returns the number of a chosen table. A dialog will open to let you choose from
        /// the found table names.
        /// </summary>
        /// <returns>the number</returns>
        public int ReadTableNames()
        {
            DataTable? dt = null;
            using ( OleDbConnection conn = OpenSource( ) )
            {
                dt = 
                    conn.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );

            }   // end: using

            if ( dt != null )
            {
                sheets = new string[ dt.Rows.Count ];
                
                for ( int i = 0; i < sheets.Length; i++ )
                {
                    sheets[ i ] = 
                        dt.Rows[ i ][ "TABLE_NAME" ].ToString()
                        ?? string.Empty;

                }

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
            DataTable? dt = null;
            using ( OleDbConnection conn = OpenSource( ) )
            {
                dt =
                    conn.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );

            }   // end: using

            if ( ( dt != null )
                && ( dt.Rows.Count > numTable ) )
            {
                return ( dt.Rows[ numTable ][ "TABLE_NAME" ].ToString() ?? string.Empty );
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
            // Spalten der Texte ( 27.11.26: vorher aus valuesDouble, falsch bzw. Absturz ohne gelesene Zahlen )
            int columns =  valuesString[0].Length;
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

    }   // end: ExcelOleDb

}   // end: namespace DbaseFrame
