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
using System.Data;
using System.Data.Common;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace DbaseFrame
{
    public class AccessOleDb
    {
        /// <summary>
        /// created on: 05.11.24
        /// last edit: 27.11.26
        /// 27.11.26: Lesen wie ExcelOleDb abgesichert: Dialog abgebrochen → keine Datei ( lastError, vorher
        ///           stillschweigend Access_Test.accdb ); OpenSource( ) meldet fehlende, gesperrte ( vorher Absturz im
        ///           ACE-Treiber ), beschädigte Datei verständlich; ohne gewählte Tabelle klare Meldung; Tabellenwahl
        ///           ohne Systemtabellen ( MSys… ); Read…List ganz oder gar nicht, neues Feld 'headers';
        ///           ReadDoubleList wandelt jede Zahlenart ( vorher Int16, Decimal, Byte, Datum still 0 ), Text → Fehler;
        ///           Tabelle ohne Datenzeile → Fehlermeldung
        /// </summary>
        Version version = new Version( "1.0.6" );
        /*
        string connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
                                 "Dbq=C:\\Path\\To\\Your\\Database.accdb;" +
                                 "Uid=Admin;Pwd=;";
        Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\path\\to\\yourdatabase.accdb;
        */
        public string sourceConnectionStart = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        public string sourceConnectionFile = "";
        public string sourceConnectionOptions = ";";
        public string sourceConnectionString = "";

        public string targetConnectionStart = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        public string targetConnectionFile = "";
        public string targetConnectionOptions = ";";
        public string targetConnectionString = "";

        public List<string[]> valuesString = new List<string[]>();
        public List<double[]> valuesDouble = new List<double[]>();
        public List<string[]> valuesTypes = new List<string[]>();

        public string[] sheets = new string[1];
        public int sheetNumber = -1;

        /// <summary>
        /// Grund, wenn keine Datei gewählt ist ( z. B. "Dateiauswahl abgebrochen" ); sonst leer ( wie ExcelOleDb ).
        /// </summary>
        public string lastError = "";

        /// <summary>
        /// Spaltennamen der zuletzt gelesenen Tabelle ( wie ExcelOleDb.headers ); nur nach fehlerfreiem Lesen gesetzt.
        /// </summary>
        public string[] headers = new string[ 0 ];

        /// <summary>
        /// Constructor for the class.
        /// Ohne Dialog und ohne Namen gilt die Demo-Datei Access_Test.accdb im Arbeitsordner ( wie bisher );
        /// Dialog abgebrochen → keine Datei, Grund in 'lastError' ( 27.11.26, wie ExcelOleDb ).
        /// </summary>
        /// <param name="file">a file name</param>
        /// <param name="silent">query for the name via dialog ?</param>
        public AccessOleDb( string file = "", bool silent = true )
        {
            sourceConnectionFile = file;

            if ( !silent && !DialogFileNameLoad( ref sourceConnectionFile ) )
            {   // abgebrochen ( 27.11.26: vorher stillschweigend Access_Test.accdb ): keine Datei
                lastError = "Dateiauswahl abgebrochen";
                sourceConnectionFile = "";
                sourceConnectionString = "";
                return;
            }   // Ende: abgebrochen
            if ( sourceConnectionFile == "" )
                sourceConnectionFile = GetDirectory() + "Access_Test.accdb";
            sourceConnectionFile = Path.GetFullPath( sourceConnectionFile );   // relativer Name → voller Pfad
            sourceConnectionString =
                sourceConnectionStart +
                sourceConnectionFile +
                sourceConnectionOptions;

        }   // end: AccessOleDb ( constructor )

        /// <summary>
        /// Öffnet die Verbindung zur Quelldatei zum Lesen ( wie ExcelOleDb.OpenSource ).
        /// </summary>
        /* Die Datei hat zu passen, sonst Fehlermeldung ( 27.11.26 ):
         *  - keine Datei gewählt ( Dialog abgebrochen ), fehlend, ganz gesperrt ( sonst Absturz im ACE-Treiber ),
         *    beschädigt / kein Access-Format / kennwortgeschützt / exklusiv geöffnet
        */
        /// <returns>die geöffnete Verbindung ( mit using verwenden )</returns>
        OleDbConnection OpenSource( )
        {
            if ( sourceConnectionFile == "" )
                throw new InvalidOperationException( "AccessOleDb: keine Datei gewählt ( " + lastError + " )." );
            if ( !File.Exists( sourceConnectionFile ) )
                throw new FileNotFoundException( $"Access-Datei nicht gefunden: {sourceConnectionFile}", sourceConnectionFile );
            ExcelOleDb.CheckReadable( sourceConnectionFile );
            OleDbConnection conn = new OleDbConnection( sourceConnectionString );
            try
            {
                conn.Open( );
            }
            catch ( OleDbException ex )
            {   // beschädigt oder kein Access-Format
                conn.Dispose( );
                throw new InvalidDataException(
                    $"Access-Datei kann nicht gelesen werden ( beschädigt, kein Access-Format, kennwortgeschützt oder exklusiv geöffnet ): {sourceConnectionFile}\n"
                    + ex.Message, ex );
            }   // Ende: beschädigt oder kein Access-Format
            return ( conn );

        }   // Ende: OleDbConnection OpenSource

        /// <summary>
        /// Name der gewählten Tabelle für das Lesen. Ohne gültige Wahl eine klare Fehlermeldung
        /// ( 27.11.26, vorher IndexOutOfRangeException; wie ExcelOleDb ).
        /// </summary>
        /// <returns>Tabellenname</returns>
        string SelectedTable( )
        {
            if ( ( sheetNumber < 0 ) || ( sheetNumber >= sheets.Length ) || string.IsNullOrEmpty( sheets[ sheetNumber ] ) )
                throw new InvalidOperationException(
                    "AccessOleDb: Keine Tabelle gewählt: zuerst ReadTableNames( ) aufrufen oder 'sheets' und 'sheetNumber' setzen." );
            return ( sheets[ sheetNumber ] );

        }   // Ende: string SelectedTable

        /// <summary>
        /// Die Tabellen des Nutzers ( TABLE_TYPE "TABLE" ), ohne Systemtabellen wie MSysACEs ( 27.11.26 ).
        /// </summary>
        /// <param name="conn">geöffnete Verbindung</param>
        /// <returns>Tabellennamen in der Reihenfolge des Treibers</returns>
        static string[] UserTables( OleDbConnection conn )
        {
            DataTable dt = conn.GetSchema( "Tables" ) ?? new DataTable( );
            List<string> namen = new List<string>( );
            foreach ( DataRow zeile in dt.Rows )
                if ( ( zeile[ "TABLE_TYPE" ]?.ToString( ) ?? "" ) == "TABLE" )
                    namen.Add( zeile[ "TABLE_NAME" ]?.ToString( ) ?? string.Empty );
            return ( namen.ToArray( ) );

        }   // Ende: string[] UserTables

        /// <summary>
        /// Spaltennamen des Lesers.
        /// </summary>
        static string[] ColumnNames( OleDbDataReader reader )
        {
            string[] namen = new string[ reader.FieldCount ];
            for ( int pos = 0; pos < reader.FieldCount; pos++ )
                namen[ pos ] = reader.GetName( pos ) ?? string.Empty;
            return ( namen );

        }   // Ende: string[] ColumnNames

        /// <summary>
        /// Reads the table's data types as anonymous array of
        /// strings into 'valuesTypes'. Ohne Prüfung auf Daten ( Untersuchung fremder Daten, wie ExcelOleDb ).
        /// </summary>
        public void ReadTypesList( )
        {
            // ganz oder gar nicht ( 27.11.26, wie ExcelOleDb ): zu Beginn leeren, lokal füllen, am Ende zuweisen
            headers = new string[ 0 ];
            valuesTypes = new List<string[]>();
            using ( OleDbConnection conn = OpenSource( ) )
            {
                OleDbCommand  command = new OleDbCommand ( $"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                string[] locHeaders = ColumnNames( reader );
                List<string[]> locTypes = new List<string[]>();

                while ( reader.Read() )
                {
                    int cols = reader.FieldCount;
                    string[] temp = new string[ cols ];
                    for ( int pos = 0; pos < cols; pos++ )
                        temp[ pos ] =
                            reader[ pos ].GetType().ToString()
                            ?? string.Empty;
                    locTypes.Add( temp );

                }
                reader.Close();
                valuesTypes = locTypes;
                headers = locHeaders;

            }   // end: using

        }   // end: ReadTypesList

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// strings into 'valuesString'. Tabelle ohne Datenzeile → InvalidDataException ( wie ExcelOleDb ).
        /// </summary>
        public void ReadStringList( )
        {
            // ganz oder gar nicht ( 27.11.26, wie ExcelOleDb ): zu Beginn leeren, lokal füllen, am Ende zuweisen
            headers = new string[ 0 ];
            valuesString = new List<string[]>();
            using ( OleDbConnection conn = OpenSource( ) )
            {
                OleDbCommand  command = new OleDbCommand ( $"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                string[] locHeaders = ColumnNames( reader );
                List<string[]> locStrings = new List<string[]>();

                while ( reader.Read() )
                {
                    int cols = reader.FieldCount;
                    string[] temp = new string[ cols ];
                    for ( int pos = 0; pos < cols; pos++ )
                        temp[ pos ] =
                            reader[ pos ].ToString( )
                            ?? string.Empty;
                    locStrings.Add( temp );

                }
                reader.Close();
                // die Daten haben zu passen: eine Tabelle ohne Datenzeile ist kein Datensatz ( wie ExcelOleDb )
                if ( locStrings.Count == 0 )
                    throw new InvalidDataException( $"AccessOleDb.ReadStringList: Tabelle '{SelectedTable( )}' enthält keine Daten." );
                valuesString = locStrings;
                headers = locHeaders;

            }   // end: using

        }   // end: ReadStringList

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// doubles into 'valuesDouble'. Umwandlung je Feld wie ExcelOleDb ( ExcelOleDb.TryCellNumber ):
        /// Zahl jeder Art ( auch Int16, Decimal, Byte ), Datum als Seriennummer, Ja/Nein 1/0, Zahl als Text,
        /// leer 0; anderer Text → InvalidDataException mit Zeile und Spalte.
        /// </summary>
        public void ReadDoubleList( )
        {
            // ganz oder gar nicht ( 27.11.26, wie ExcelOleDb ): zu Beginn leeren, lokal füllen, am Ende zuweisen
            headers = new string[ 0 ];
            valuesDouble = new List<double[]>();
            using ( OleDbConnection conn = OpenSource( ) )
            {
                OleDbCommand command = new OleDbCommand($"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                string[] locHeaders = ColumnNames( reader );
                List<double[]> locDoubles = new List<double[]>();

                while ( reader.Read() )
                {
                    int cols = reader.FieldCount;
                    double[] temp = new double[ cols ];
                    for ( int pos = 0; pos < cols; pos++ )
                    {   // jeder Wert einzeln geprüft ( 27.11.26: vorher nur Int32, Double, Single, alles andere still 0 )
                        if ( !ExcelOleDb.TryCellNumber( reader[ pos ], out temp[ pos ] ) )
                            throw new InvalidDataException( $"AccessOleDb.ReadDoubleList: Zeile {locDoubles.Count + 1} der Tabelle, "
                                + $"Spalte '{locHeaders[ pos ]}' enthält den Text '{reader[ pos ]}' statt einer Zahl ( nur leere Felder werden 0 )." );
                    }   // Ende: jeder Wert einzeln geprüft
                    locDoubles.Add( temp );

                }
                reader.Close();
                // die Daten haben zu passen: eine Tabelle ohne Datenzeile ist kein Datensatz ( wie ExcelOleDb )
                if ( locDoubles.Count == 0 )
                    throw new InvalidDataException( $"AccessOleDb.ReadDoubleList: Tabelle '{SelectedTable( )}' enthält keine Daten." );
                valuesDouble = locDoubles;
                headers = locHeaders;

            }   // end: using

        }   // end: ReadDoubleList

        /// <summary>
        /// Returns the number of a chosen table. A dialog will open to let you choose from
        /// the found table names ( nur Tabellen des Nutzers, ohne Systemtabellen, 27.11.26 ).
        /// </summary>
        /// <returns>the number, -1 = keine gewählt</returns>
        public int ReadTableNames( )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                sheets = UserTables( conn );
            }   // end: using

            DialogTablesChoice choice = new DialogTablesChoice( sheets );
            sheetNumber = choice.index;
            return ( sheetNumber );

        }   // end: ReadTableNames

        /// <summary>
        /// Direct query for the table name ( Zählung wie ReadTableNames, ohne Systemtabellen ).
        /// </summary>
        /// <param name="numTable">number of the sheet</param>
        /// <returns>the name or 'string.empty'</returns>
        public string GetTableName( int numTable )
        {
            string[] tabellen;
            using ( OleDbConnection conn = OpenSource( ) )
            {
                tabellen = UserTables( conn );
            }   // end: using

            if ( ( numTable >= 0 ) && ( numTable < tabellen.Length ) )
                return ( tabellen[ numTable ] );
            return ( string.Empty );

        }   // end: GetTableName

        public void ColumnNameDummy()
        {
            using ( OleDbConnection conn = new OleDbConnection( sourceConnectionString ) )
            {
                conn.Open();
                
                // Get the schema table for columns
                
                DataTable schemaTable = 
                    conn.GetSchema("Columns")
                    ?? new DataTable();

                // Iterate through the schema table rows to access column headers (field names)
                foreach ( DataColumn column in schemaTable.Columns )
                {
                    string columnName = column.ColumnName;
                    // Use the column name as needed
                }
            }
        }

        /// <summary>
        /// Target file name for the writing is chosen. Produces the
        /// 'targetConnectionString' for convenience.
        /// </summary>
        /// <param name="file">already known ?</param>
        /// <param name="silent">use the dialog ?</param>
        public void ChooseTarget( ref string file, bool silent = true )
        {
            targetConnectionFile = file;

            bool ok = false;
            if ( !silent )
                ok = DialogFileNameSave( ref targetConnectionFile );
            if ( targetConnectionFile != "" )
            {
                targetConnectionString =
                    targetConnectionStart +
                    targetConnectionFile +
                    targetConnectionOptions;
                file = targetConnectionFile;
            }
            else
            {
                targetConnectionString =
                    sourceConnectionString +
                    GetDirectory() +
                    "newAccess_Test.accdb" +
                    targetConnectionOptions;
                file = GetDirectory() + "newAccess_Test.accdb";
            }
            targetConnectionFile = file;
            //Message.Show( file );

        }   // end: ChooseTarget


        /// <summary>
        /// Intern data list double will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        public void WriteListDoubleToNewTarget( string newFileTarget = "", string newTableName = "newDoubles" )
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
                foreach ( double[] row in valuesDouble )
                {
                    command.CommandText = commandInsert;
                    command.Parameters.Clear();

                    for ( int pos = 0; pos < row.Length; pos++ )
                        command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );

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
                        command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }

        }   // end: WriteListStringToNewTarget

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
            dialog.DefaultExt = ".accdb"; // Default file extension
            dialog.Filter = "Access save file (.accdb)|*.accdb"; // Filter files by extension
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
            fileName = string.Empty;
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
            dialog.DefaultExt = ".accdb"; // Default file extension
            dialog.Filter = "Access save file (.accdb)|*.accdb"; // Filter files by extension
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



    }   // end: public class AccessOleDb

}   // end: namespace DbaseFrame

