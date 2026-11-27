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
        ///           Write…ToNewTarget: Dialog abgebrochen → nichts schreiben ( vorher NewTarget.xlsx );
        ///           WriteListStringToNewTarget zählt die Spalten der Texte ( vorher der Zahlen ) und prüft vorher die
        ///           Textlänge ( höchstens 255 Zeichen, sonst Fehlermeldung, nichts geschrieben );
        ///           Dateidialog abgebrochen → keine Datei ( 'lastError', wie ExcelInterop ), vorher stillschweigend
        ///           Parable_Demo.xlsx aus dem Arbeitsordner; WriteListDoubleToNewTarget ohne die zwei MessageBoxen
        ///           mit dem SQL-Text ( Testausgabe, wie schon in WriteListStringToNewTarget );
        ///           Write…ToNewTarget liefern bool ( geschrieben / nichts geschrieben ); neues Feld 'headers'
        ///           ( Spaltennamen beim Lesen ), bei useHeader Spaltennamen der neuen Tabelle ( vorher 0, 1, 2, ... );
        ///           ganz gesperrte Datei → IOException ( CheckReadable ), vorher Absturz im ACE-Treiber;
        ///           Read…List leeren ihre Liste zu Beginn ( nach einem Fehler keine Daten eines früheren Lesens )
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

        /// <summary>
        /// Grund, wenn keine Datei gewählt ist ( z. B. "Dateiauswahl abgebrochen" ); sonst leer ( wie ExcelInterop ).
        /// </summary>
        public string lastError = "";
        public string targetFileName = "";
        public List<string[]> valuesString = new List<string[]>();
        public List<double[]> valuesDouble = new List<double[]>();
        public List<string[]> valuesTypes = new List<string[]>();

        /// <summary>
        /// Spaltennamen der zuletzt gelesenen Tabelle ( wie ExcelInterop.headers ): bei useHeader die Kopfzeile,
        /// sonst die Namen des Treibers ( F1, F2, ... ). Nur nach fehlerfreiem Lesen gesetzt, sonst leer.
        /// Die Schreibfunktionen nehmen sie bei useHeader als Spaltennamen der neuen Tabelle ( 27.11.26 ).
        /// </summary>
        public string[] headers = new string[ 0 ];
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
            {   // Datei per Dialog
                ok = DialogFileNameLoad( ref fileName );
                if ( !ok )
                {   // abgebrochen ( Vorgabe vom 27.11.26, wie ExcelInterop ): keine Datei, keine Ersatzdatei
                    lastError = "Dateiauswahl abgebrochen";
                    fileName = "";
                    sourceFile = "";
                    connectionString = "";
                    return;
                }   // Ende: abgebrochen
            }   // Ende: Datei per Dialog
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
            // keine Datei gewählt ( Dialog abgebrochen ): klare Meldung statt "nicht gefunden: " ohne Namen
            if ( sourceFile == "" )
                throw new InvalidOperationException( "ExcelOleDb: keine Datei gewählt ( " + lastError + " )." );
            if ( !File.Exists( sourceFile ) )
                throw new FileNotFoundException( $"Excel-Datei nicht gefunden: {sourceFile}", sourceFile );
            CheckReadable( sourceFile );
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
        /// Prüft vor dem Öffnen mit dem ACE-Treiber, ob sich die Datei lesen lässt.
        /// </summary>
        /* Grund ( 27.11.26, gemessen ): Hält ein anderer Prozess die Datei ohne Lesefreigabe offen
         * ( FileShare.None ), stürzt der ACE-Treiber in OleDbConnection.Open( ) mit einer Zugriffsverletzung
         * ( 0xC0000005 ) ab; das beendet das ganze Programm und lässt sich nicht abfangen.
         * Eine normal in Excel geöffnete Datei besteht die Prüfung; dafür meldet sich der Treiber selbst
         * ( "bereits exklusiv geöffnet", siehe OpenSource ).
        */
        /// <param name="datei">voller Dateiname</param>
        /// <exception cref="IOException">Datei gesperrt oder ohne Leserecht</exception>
        internal static void CheckReadable( string datei )
        {
            try
            {
                using ( FileStream probe = new FileStream( datei, FileMode.Open, FileAccess.Read, FileShare.ReadWrite ) )
                {   // nur öffnen und gleich schließen
                }   // Ende: nur öffnen und gleich schließen
            }
            catch ( Exception ex ) when ( ex is IOException || ex is UnauthorizedAccessException )
            {   // gesperrt: verständliche Meldung statt Absturz im Treiber
                throw new IOException( $"Datei ist gesperrt oder nicht lesbar ( woanders ohne Lesefreigabe geöffnet? ): {datei}\n"
                    + ex.Message, ex );
            }   // Ende: gesperrt

        }   // Ende: void CheckReadable

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

        /// <summary>
        /// Höchstlänge eines Textes beim Schreiben: Spaltentyp VARCHAR des ACE-Treibers ( mit Excel geprüft 27.11.26 )
        /// </summary>
        public const int maxTextLength = 255;

        /// <summary>
        /// Prüft Texte vor dem Schreiben auf die Höchstlänge 'maxTextLength'.
        /// </summary>
        /// <param name="rows">die Zeilen</param>
        /// <param name="funktion">aufrufende Funktion für die Meldung</param>
        /// <exception cref="ArgumentException">ein Text ist zu lang, mit Position ( spalte, zeile ) ab 0</exception>
        static void CheckTextLength( List<string[]> rows, string funktion )
        {
            for ( int y = 0; y < rows.Count; y++ )
                for ( int x = 0; x < rows[ y ].Length; x++ )
                    if ( ( rows[ y ][ x ]?.Length ?? 0 ) > maxTextLength )
                        throw new ArgumentException( "ExcelOleDb." + funktion + ": Text an ( " + x + ", " + y
                            + " ) hat " + rows[ y ][ x ].Length + " Zeichen, höchstens " + maxTextLength
                            + " sind möglich. Nichts geschrieben." );

        }   // Ende: void CheckTextLength

        // --------------------------------------------     the routines

        /// <summary>
        /// Reads the table's data types as anonymous array of
        /// strings into 'valuesTypes'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadTypesList( )
        {
            // ganz oder gar nicht ( 27.11.26, wie ExcelInterop ): zu Beginn leeren, lokal füllen, am Ende zuweisen
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
                conn.Close();
                valuesTypes = locTypes;
                headers = locHeaders;

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
            // ganz oder gar nicht ( 27.11.26, wie ExcelInterop ): zu Beginn leeren, lokal füllen, am Ende zuweisen
            headers = new string[ 0 ];
            valuesString = new List<string[]>();
            using ( OleDbConnection conn = OpenSource( ) )
            {
                OleDbCommand command = new OleDbCommand( $"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                string[] locHeaders = ColumnNames( reader );

                List<string[]> locStrings = new List<string[]>();

                while ( reader.Read() )
                {
                    string[] temp = new string[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        temp[ pos ] =
                            reader[ pos ].ToString()
                            ?? string.Empty;
                    locStrings.Add( temp );

                }
                valuesString = locStrings;
                headers = locHeaders;

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
            // ganz oder gar nicht ( 27.11.26, wie ExcelInterop ): zu Beginn leeren, lokal füllen, am Ende zuweisen
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
                    double[] temp = new double[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        if ( reader[ pos ].GetType() == typeof( double ) )
                            temp[ pos ] = reader.GetDouble( pos );
                    locDoubles.Add( temp );

                }
                valuesDouble = locDoubles;
                headers = locHeaders;

            }   // end: using

        }   // end: ReadDoubleList

        /// <summary>
        /// Spaltennamen des Lesers ( bei HDR=YES die Kopfzeile, sonst F1, F2, ... des Treibers ).
        /// </summary>
        /// <param name="reader">geöffneter Leser</param>
        /// <returns>je Spalte ein Name</returns>
        static string[] ColumnNames( OleDbDataReader reader )
        {
            string[] namen = new string[ reader.FieldCount ];
            for ( int pos = 0; pos < reader.FieldCount; pos++ )
                namen[ pos ] = reader.GetName( pos ) ?? string.Empty;
            return ( namen );

        }   // Ende: string[] ColumnNames

        /// <summary>
        /// Spaltennamen für eine neue Tabelle: bei useHeader und passender Spaltenzahl die gelesenen 'headers'
        /// ( Kopfzeile bleibt erhalten, 27.11.26 ), sonst wie bisher 0, 1, 2, ...
        /// Für ACE bereinigt: . ! [ ] ` werden _, leer → Spaltennummer, doppelt → Anhang _2, _3, ...
        /// </summary>
        /// <param name="columns">Spaltenzahl der Daten</param>
        /// <returns>je Spalte ein eindeutiger Name</returns>
        string[] ColumnNamesForWrite( int columns )
        {
            bool kopf = useHeader && ( headers.Length == columns );
            string[] namen = new string[ columns ];
            HashSet<string> vergeben = new HashSet<string>( StringComparer.OrdinalIgnoreCase );
            for ( int pos = 0; pos < columns; pos++ )
            {   // je Spalte
                string name = kopf ? headers[ pos ] : pos.ToString( );
                foreach ( char zeichen in ".![]`" )
                    name = name.Replace( zeichen, '_' );
                name = name.Trim( );
                if ( name.Length > 64 )
                    name = name.Substring( 0, 64 );
                if ( name == "" )
                    name = pos.ToString( );
                string basis = name;
                for ( int anhang = 2; !vergeben.Add( name ); anhang++ )
                    name = basis + "_" + anhang;
                namen[ pos ] = name;
            }   // Ende: je Spalte
            return ( namen );

        }   // Ende: string[] ColumnNamesForWrite

        /// <summary>
        /// Baut CREATE TABLE und INSERT INTO für eine neue Tabelle mit den Spaltennamen aus ColumnNamesForWrite.
        /// </summary>
        /// <param name="tabelle">Tabellenname</param>
        /// <param name="columns">Spaltenzahl</param>
        /// <param name="typ">SQL-Typ der Spalten ( DOUBLE, VARCHAR )</param>
        /// <param name="commandInsert">INSERT INTO mit Parametern @0, @1, ...</param>
        /// <returns>CREATE TABLE</returns>
        string TableCommands( string tabelle, int columns, string typ, out string commandInsert )
        {
            string[] namen = ColumnNamesForWrite( columns );
            List<string> definitionen = new List<string>( );
            List<string> spalten = new List<string>( );
            List<string> parameter = new List<string>( );
            for ( int pos = 0; pos < columns; pos++ )
            {   // je Spalte
                definitionen.Add( $"[{namen[ pos ]}] {typ}" );
                spalten.Add( $"[{namen[ pos ]}]" );
                parameter.Add( $"@{pos}" );
            }   // Ende: je Spalte
            commandInsert = $"INSERT INTO [{tabelle}] ( {string.Join( ", ", spalten )} ) VALUES ( {string.Join( ", ", parameter )} );";
            return ( $"CREATE TABLE [{tabelle}] ( {string.Join( ", ", definitionen )} );" );

        }   // Ende: string TableCommands

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
        /// <returns>false, wenn der Dialog abgebrochen wurde: dann nichts schreiben</returns>
        public bool ChooseTarget( ref string file, bool silent = true )
        {
            targetFileName = file;

            if ( !silent && !DialogFileNameSave( ref targetFileName ) )
            {   // Dialog abgebrochen ( 27.11.26: vorher wurde dann NewTarget.xlsx im Programmordner beschrieben )
                return ( false );
            }   // Ende: Dialog abgebrochen
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
            return ( true );

        }   // end: ChooseTarget


        /// <summary>
        /// Intern data list double will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        /// <returns>true = geschrieben; false = keine Daten oder Dialog abgebrochen, nichts geschrieben
        /// ( 27.11.26, wie ExcelInterop )</returns>
        public bool WriteListDoubleToNewTarget( string newFileTarget = "",string newTableName = "newDoubles" )
        {
            if ( ( valuesDouble.Count < 1 ) || ( valuesDouble[ 0 ].Length == 0 ) )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return ( false );
            }

            bool overwrite = false;
            while ( !overwrite )
            {
                // ohne Namen per Dialog; abgebrochen: nichts schreiben
                if ( !ChooseTarget( ref newFileTarget, newFileTarget != "" ) )
                    return ( false );

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
            // craft the 'CREATE TABLE' and 'INSERT INTO'; Spaltennamen bei useHeader aus der gelesenen
            // Kopfzeile ( 27.11.26: vorher immer 0, 1, 2, ... )
            int columns =  valuesDouble[0].Length;
            string commandCreate = TableCommands( newTableName, columns, "DOUBLE", out string commandInsert );
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
            return ( true );

        }   // end: WriteListDoubleToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        /// <returns>true = geschrieben; false = keine Daten oder Dialog abgebrochen, nichts geschrieben
        /// ( 27.11.26, wie ExcelInterop )</returns>
        public bool WriteListStringToNewTarget( string newFileTarget = "", string newTableName = "newStrings" )
        {
            if ( ( valuesString.Count < 1 ) || ( valuesString[ 0 ].Length == 0 ) )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return ( false );
            }
            // die Daten haben zu passen: vor Dialog und Datei prüfen ( vorher Abbruch mitten im Schreiben,
            // die Zieldatei blieb mit leerer Tabelle zurück )
            CheckTextLength( valuesString, "WriteListStringToNewTarget" );

            bool overwrite = false;
            while ( !overwrite )
            {
                // ohne Namen per Dialog; abgebrochen: nichts schreiben
                if ( !ChooseTarget( ref newFileTarget, newFileTarget != "" ) )
                    return ( false );

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
            // Spalten der Texte ( 27.11.26: vorher aus valuesDouble, falsch bzw. Absturz ohne gelesene Zahlen );
            // Spaltennamen bei useHeader aus der gelesenen Kopfzeile ( 27.11.26: vorher immer 0, 1, 2, ... )
            int columns =  valuesString[0].Length;
            string commandCreate = TableCommands( newTableName, columns, "VARCHAR", out string commandInsert );
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
            return ( true );

        }   // end: WriteListStringToNewTarget

    }   // end: ExcelOleDb

}   // end: namespace DbaseFrame
