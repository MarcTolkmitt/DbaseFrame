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
using System.Runtime.InteropServices;
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
        ///           Tabelle ohne Datenzeile → Fehlermeldung;
        ///           Schreiben neu ( Entscheidung des Nutzers „mit ADOX“ ): neue Datei per ADOX ( OleDb kann keine anlegen,
        ///           vorher scheiterte jeder Aufruf ), Write…ToNewTarget liefern bool, ChooseTarget mit Abbruch, Spaltennamen
        ///           aus 'headers', Spaltenzahl der Texte richtig, Textlänge vorher geprüft, Verbindung bleibt offen
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
        /// Ohne Namen und ohne Dialog: newAccess_Test.accdb im Arbeitsordner ( wie bisher gemeint; 27.11.26: vorher
        /// wurde dafür fälschlich die Quell-Verbindung vorangestellt ). Name wird vervollständigt ( .accdb, voller Pfad ).
        /// </summary>
        /// <param name="file">already known ?</param>
        /// <param name="silent">use the dialog ?</param>
        /// <returns>false, wenn der Dialog abgebrochen wurde: dann nichts schreiben ( 27.11.26, wie ExcelOleDb )</returns>
        public bool ChooseTarget( ref string file, bool silent = true )
        {
            targetConnectionFile = file;

            if ( !silent && !DialogFileNameSave( ref targetConnectionFile ) )
                return ( false );   // abgebrochen
            if ( targetConnectionFile == "" )
                targetConnectionFile = GetDirectory() + "newAccess_Test.accdb";
            if ( Path.GetExtension( targetConnectionFile ) == "" )
                targetConnectionFile += ".accdb";
            targetConnectionFile = Path.GetFullPath( targetConnectionFile );
            targetConnectionString =
                targetConnectionStart +
                targetConnectionFile +
                targetConnectionOptions;
            file = targetConnectionFile;
            return ( true );

        }   // end: ChooseTarget

        /// <summary>
        /// Legt eine neue, leere Access-Datei an ( 27.11.26 ). Der ACE-Treiber kann das über OleDb nicht
        /// ( "Datei nicht gefunden" ), daher über ADOX ( COM "ADOX.Catalog", kommt mit dem ACE-Treiber ).
        /// </summary>
        /// <param name="datei">voller Dateiname ( .accdb )</param>
        /// <exception cref="InvalidOperationException">ADOX nicht vorhanden ( ACE-Treiber fehlt )</exception>
        void CreateDatabase( string datei )
        {
            Type katalogTyp = Type.GetTypeFromProgID( "ADOX.Catalog" )
                ?? throw new InvalidOperationException(
                    "AccessOleDb: ADOX fehlt, eine neue Access-Datei kann nicht angelegt werden ( ACE-Treiber installiert? )." );
            dynamic katalog = Activator.CreateInstance( katalogTyp )!;
            try
            {   // anlegen; die von ADOX geöffnete Verbindung sofort schließen, sonst bleibt die Datei gesperrt
                dynamic verbindung = katalog.Create( targetConnectionStart + datei + targetConnectionOptions );
                verbindung.Close( );
                Marshal.ReleaseComObject( verbindung );
            }   // Ende: anlegen
            finally
            {   // COM-Objekt freigeben ( Grundsatz: jedes festgehaltene COM-Objekt freigeben )
                Marshal.ReleaseComObject( katalog );
            }   // Ende: COM-Objekt freigeben

        }   // Ende: void CreateDatabase

        /// <summary>
        /// Gemeinsamer Ablauf beider Schreibfunktionen ( 27.11.26 ): Ziel wählen ( Abbruch → false ), vorhandene Datei nur
        /// nach Rückfrage ersetzen, neue Datei per ADOX anlegen, Tabelle anlegen und Zeilen einfügen. Die Verbindung bleibt
        /// zwischen CREATE und INSERT offen ( Lehre aus ExcelOleDb ). Spaltennamen aus 'headers' ( ExcelOleDb.ColumnNamesForWrite ).
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Tabellenname</param>
        /// <param name="columns">Spaltenzahl der Daten</param>
        /// <param name="typ">SQL-Typ ( DOUBLE, VARCHAR )</param>
        /// <param name="parameterTyp">OleDb-Typ der Werte</param>
        /// <param name="zeilen">die Zeilen als Werte</param>
        /// <returns>true = geschrieben, false = Dialog abgebrochen</returns>
        bool WriteTable( string newFileTarget, string newTableName, int columns, string typ, OleDbType parameterTyp,
            IEnumerable<object[]> zeilen )
        {
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

            CreateDatabase( newFileTarget );
            string commandCreate = ExcelOleDb.TableCommands( newTableName,
                ExcelOleDb.ColumnNamesForWrite( headers, true, columns ), typ, out string commandInsert );
            using ( OleDbConnection connection = new OleDbConnection( targetConnectionString ) )
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand( commandCreate, connection );
                command.ExecuteNonQuery();
                foreach ( object[] zeile in zeilen )
                {   // je Zeile
                    command.CommandText = commandInsert;
                    command.Parameters.Clear();
                    for ( int pos = 0; pos < zeile.Length; pos++ )
                        command.Parameters.Add( $"@{pos}", parameterTyp ).Value = zeile[ pos ] ?? DBNull.Value;
                    command.ExecuteNonQuery();
                }   // Ende: je Zeile
                connection.Close();
            }
            return ( true );

        }   // Ende: bool WriteTable

        /// <summary>
        /// Intern data list double will be written into a new
        /// Access file. If not given a name a dialog will query for it ( 27.11.26: neue Datei per ADOX,
        /// vorher scheiterte jeder Aufruf mit "Datei nicht gefunden" ).
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Tabellenname</param>
        /// <returns>true = geschrieben; false = keine Daten oder Dialog abgebrochen, nichts geschrieben</returns>
        public bool WriteListDoubleToNewTarget( string newFileTarget = "", string newTableName = "newDoubles" )
        {
            if ( ( valuesDouble.Count < 1 ) || ( valuesDouble[ 0 ].Length == 0 ) )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return ( false );
            }
            List<object[]> zeilen = new List<object[]>( );
            foreach ( double[] zeile in valuesDouble )
                zeilen.Add( zeile.Cast<object>( ).ToArray( ) );
            return ( WriteTable( newFileTarget, newTableName, valuesDouble[ 0 ].Length, "DOUBLE", OleDbType.Double, zeilen ) );

        }   // end: WriteListDoubleToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new
        /// Access file. If not given a name a dialog will query for it. Texte höchstens 255 Zeichen
        /// ( VARCHAR, wie ExcelOleDb ), vorher geprüft, sonst nichts geschrieben.
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Tabellenname</param>
        /// <returns>true = geschrieben; false = keine Daten oder Dialog abgebrochen, nichts geschrieben</returns>
        public bool WriteListStringToNewTarget( string newFileTarget = "", string newTableName = "newStrings" )
        {
            if ( ( valuesString.Count < 1 ) || ( valuesString[ 0 ].Length == 0 ) )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return ( false );
            }
            // die Daten haben zu passen: vor Dialog und Datei prüfen ( wie ExcelOleDb )
            ExcelOleDb.CheckTextLength( valuesString, "AccessOleDb.WriteListStringToNewTarget" );
            List<object[]> zeilen = new List<object[]>( );
            foreach ( string[] zeile in valuesString )
                zeilen.Add( zeile.Cast<object>( ).ToArray( ) );
            // Spalten der Texte ( 27.11.26: vorher aus valuesDouble, falsch bzw. Absturz ohne gelesene Zahlen )
            return ( WriteTable( newFileTarget, newTableName, valuesString[ 0 ].Length, "VARCHAR", OleDbType.VarChar, zeilen ) );

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

