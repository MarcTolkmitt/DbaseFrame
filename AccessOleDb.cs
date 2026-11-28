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
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DbaseFrame
{
    /// <summary>
    /// Reads and writes Access databases ( .accdb ) through the ACE OleDb driver: table names, column names, texts,
    /// numbers and cell types of a table; writes one new table into a new database. History: see the field 'version'.
    /// </summary>
    public class AccessOleDb
    {
        /// <summary>
        /// created on: 05.11.24
        /// last edit: 28.11.26
        /// 27.11.26: Lesen wie ExcelOleDb abgesichert: Dialog abgebrochen → keine Datei ( lastError, vorher
        ///           stillschweigend Access_Test.accdb ); OpenSource( ) meldet fehlende, gesperrte ( vorher Absturz im
        ///           ACE-Treiber ), beschädigte Datei verständlich; ohne gewählte Tabelle klare Meldung; Tabellenwahl
        ///           ohne Systemtabellen ( MSys… ); Read…List ganz oder gar nicht, neues Feld 'headers';
        ///           ReadDoubleList wandelt jede Zahlenart ( vorher Int16, Decimal, Byte, Datum still 0 ), Text → Fehler;
        ///           Tabelle ohne Datenzeile → Fehlermeldung;
        ///           Schreiben neu ( Entscheidung des Nutzers „mit ADOX“ ): neue Datei per ADOX ( OleDb kann keine anlegen,
        ///           vorher scheiterte jeder Aufruf ), Write…ToNewTarget liefern bool, ChooseTarget mit Abbruch, Spaltennamen
        ///           aus 'headers', Spaltenzahl der Texte richtig, Textlänge vorher geprüft, Verbindung bleibt offen;
        ///           ohne System.Data.Odbc ( wurde nicht mehr gebraucht, Paket entfernt )
        /// 28.11.26: ReadColumnNames( numTable ) ersetzt ColumnNameDummy ( las nur die Spalten der Schematabelle und
        ///           verwarf sie; bleibt als leere [Obsolete]-Methode eine Version lang stehen ); gleich in ExcelOleDb
        ///           und ExcelInterop;
        ///           Punkt l: keine MessageBox mehr aus der Bibliothek: keine Daten → InvalidOperationException,
        ///           vorhandene Zieldatei nur nach Bestätigung im Dialog oder mit overwrite: true, sonst IOException
        ///           ( ExcelOleDb.CheckHasData / PrepareTarget ); Rückgabe false heißt nur noch „Dialog abgebrochen“;
        ///           Befund H: neue Datei als Zwischendatei, die alte erst danach ersetzt ( ExcelOleDb.CommitTarget );
        ///           Tabellenname vor dem Dialog geprüft ( ExcelOleDb.CheckTableName ), CREATE TABLE übersetzt;
        ///           neues Feld writtenTableName: Name der geschriebenen Tabelle, wie GetTableName / GetSheetName ihn liefert
        ///           28.11.26 ( 2 ): GetTableName wirft bei ungültiger Nummer ArgumentOutOfRangeException ( vorher leer ), neu GetTableCount
        /// </summary>
        Version version = new Version( "1.0.7" );
        /*
        string connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
                                 "Dbq=C:\\Path\\To\\Your\\Database.accdb;" +
                                 "Uid=Admin;Pwd=;";
        Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\path\\to\\yourdatabase.accdb;
        */
        /// <summary>
        /// First part of the source connection string: provider and "Data Source=", followed by the file name.
        /// </summary>
        public string sourceConnectionStart = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        /// <summary>
        /// Full path of the source database; empty = no file chosen ( see 'lastError' ).
        /// </summary>
        public string sourceConnectionFile = "";
        /// <summary>
        /// Part of the source connection string after the file name.
        /// </summary>
        public string sourceConnectionOptions = ";";
        /// <summary>
        /// Connection string of the source: sourceConnectionStart + sourceConnectionFile + sourceConnectionOptions.
        /// </summary>
        public string sourceConnectionString = "";

        /// <summary>
        /// First part of the target connection string: provider and "Data Source=", followed by the file name.
        /// </summary>
        public string targetConnectionStart = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        /// <summary>
        /// Full path of the target database chosen by ChooseTarget ( called by Write…ToNewTarget ).
        /// </summary>
        public string targetConnectionFile = "";
        /// <summary>
        /// Part of the target connection string after the file name.
        /// </summary>
        public string targetConnectionOptions = ";";
        /// <summary>
        /// Connection string of the target database chosen by ChooseTarget.
        /// </summary>
        public string targetConnectionString = "";

        /// <summary>
        /// Name der zuletzt geschriebenen Tabelle, so wie GetTableName ihn liefert ( 28.11.26, wie ExcelOleDb ). Der
        /// Access-Treiber übernimmt den Namen unverändert ( gemessen ). Zu Beginn jedes Schreibens leer, nur nach
        /// gelungenem Schreiben gesetzt.
        /// </summary>
        public string writtenTableName = "";

        /// <summary>
        /// Texts of the last ReadStringList, one array per row; also the data for WriteListStringToNewTarget.
        /// </summary>
        public List<string[]> valuesString = new List<string[]>();
        /// <summary>
        /// Numbers of the last ReadDoubleList, one array per row; also the data for WriteListDoubleToNewTarget.
        /// </summary>
        public List<double[]> valuesDouble = new List<double[]>();
        /// <summary>
        /// .NET type name of every cell of the last ReadTypesList, one array per row.
        /// </summary>
        public List<string[]> valuesTypes = new List<string[]>();

        /// <summary>
        /// Names of the user tables of the source database, filled by ReadTableNames.
        /// </summary>
        public string[] sheets = new string[1];
        /// <summary>
        /// Index of the chosen table in 'sheets'; -1 = none chosen yet.
        /// </summary>
        public int sheetNumber = -1;

        /// <summary>
        /// Grund, wenn keine Datei gewählt ist ( z. B. "Dateiauswahl abgebrochen" ); sonst leer ( wie ExcelOleDb ).
        /// </summary>
        public string lastError = "";

        /// <summary>
        /// Spaltennamen der zuletzt gelesenen Tabelle ( wie ExcelOleDb.headers ); nur nach fehlerfreiem Lesen gesetzt.
        /// </summary>
        public string[] headers = new string[ 0 ];

        // Spaltennamen je Datensatz ( 27.11.26, wie ExcelOleDb ): gesetzt zusammen mit valuesString bzw. valuesDouble
        string[] headersString = new string[ 0 ];
        string[] headersDouble = new string[ 0 ];

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

        /* Die Datei hat zu passen, sonst Fehlermeldung ( 27.11.26 ):
         *  - keine Datei gewählt ( Dialog abgebrochen ), fehlend, ganz gesperrt ( sonst Absturz im ACE-Treiber ),
         *    beschädigt / kein Access-Format / kennwortgeschützt / exklusiv geöffnet
        */
        /// <summary>
        /// Öffnet die Verbindung zur Quelldatei zum Lesen ( wie ExcelOleDb.OpenSource ).
        /// </summary>
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
            headersString = new string[ 0 ];
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
                headersString = locHeaders;
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
            headersDouble = new string[ 0 ];
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
                headersDouble = locHeaders;
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

#if NETSTANDARD2_0
            sheetNumber = -1;   // netstandard2.0 ( 28.11.26 ): kein Auswahldialog; Tabelle über sheetNumber bzw. SelectSheet wählen
#else
            DialogTablesChoice choice = new DialogTablesChoice( sheets );
            sheetNumber = choice.index;
#endif
            return ( sheetNumber );

        }   // end: ReadTableNames

        /// <summary>
        /// Direct query for the table name ( Zählung wie ReadTableNames, ohne Systemtabellen ).
        /// </summary>
        /// <param name="numTable">number of the table, 0 = first; valid up to GetTableCount( ) - 1</param>
        /// <returns>the name</returns>
        /// <exception cref="ArgumentOutOfRangeException">no table with this number ( since 1.0.9; before: empty text )</exception>
        public string GetTableName( int numTable )
        {
            string[] tabellen;
            using ( OleDbConnection conn = OpenSource( ) )
            {
                tabellen = UserTables( conn );
            }   // end: using

            // 28.11.26: ungültige Nummer → Ausnahme wie ReadColumnNames ( Entscheidung „Angleichen“; vorher leerer Text )
            if ( ( numTable < 0 ) || ( numTable >= tabellen.Length ) )
                throw new ArgumentOutOfRangeException( nameof( numTable ), numTable,
                    $"AccessOleDb.GetTableName: keine Tabelle Nr. {numTable} ( vorhanden: {tabellen.Length}, gezählt ab 0 )." );
            return ( tabellen[ numTable ] );

        }   // end: GetTableName

        /// <summary>
        /// Number of user tables in the source database, counted like GetTableName / ReadColumnNames.
        /// </summary>
        /// <returns>the number of tables</returns>
        public int GetTableCount( )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                return ( UserTables( conn ).Length );
            }   // end: using

        }   // end: int GetTableCount

        /// <summary>
        /// Spaltennamen einer Tabelle, ohne Daten zu lesen ( ersetzt ColumnNameDummy, 28.11.26 ).
        /// Zählung der Tabellen wie ReadTableNames / GetTableName ( ohne Systemtabellen ).
        /// Die Namen sind dieselben, die ein Read…List in 'headers' liefert.
        /// </summary>
        /// <param name="numTable">Nummer der Tabelle, 0 = erste</param>
        /// <returns>die Spaltennamen</returns>
        /// <exception cref="ArgumentOutOfRangeException">keine Tabelle mit dieser Nummer</exception>
        public string[] ReadColumnNames( int numTable )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                string[] tabellen = UserTables( conn );
                if ( ( numTable < 0 ) || ( numTable >= tabellen.Length ) )
                    throw new ArgumentOutOfRangeException( nameof( numTable ),
                        $"AccessOleDb: Tabelle Nr. {numTable} gibt es nicht ( {tabellen.Length} Tabellen in {sourceConnectionFile} )." );

                using ( OleDbCommand command = new OleDbCommand( $"SELECT * FROM [{tabellen[ numTable ]}]", conn ) )
                using ( OleDbDataReader reader = command.ExecuteReader( CommandBehavior.SchemaOnly ) )
                {   // nur das Schema, keine Datenzeile
                    return ( ColumnNames( reader ) );
                }   // Ende: nur das Schema, keine Datenzeile
            }   // Ende: using

        }   // Ende: string[] ReadColumnNames

        /// <summary>
        /// Veraltet ( 28.11.26 ): tat nichts ( las die Spalten der Schematabelle und verwarf sie ). Bleibt eine Version lang
        /// stehen, damit Aufrufer aus den Paketen bis 1.0.8 weiter übersetzen; stattdessen ReadColumnNames( numTable ).
        /// </summary>
        [Obsolete( "ColumnNameDummy tat nichts und entfällt in einer späteren Version: ReadColumnNames( int numTable ) verwenden." )]
        public void ColumnNameDummy( )
        {
        }   // Ende: void ColumnNameDummy

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
        /// nach Bestätigung im Dialog oder mit 'overwrite' ersetzen ( 28.11.26, ExcelOleDb.PrepareTarget statt MessageBox ),
        /// neue Datei per ADOX anlegen, Tabelle anlegen und Zeilen einfügen. Die Verbindung bleibt
        /// zwischen CREATE und INSERT offen ( Lehre aus ExcelOleDb ). Spaltennamen aus 'headers' ( ExcelOleDb.ColumnNamesForWrite ).
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Tabellenname: höchstens 64 Zeichen, ohne . ! ` [ ], kein führendes Leerzeichen</param>
        /// <param name="overwrite">true = eine vorhandene Zieldatei ohne Rückfrage ersetzen</param>
        /// <param name="funktion">aufrufende Funktion für Meldungen</param>
        /// <param name="columns">Spaltenzahl der Daten</param>
        /// <param name="typ">SQL-Typ ( DOUBLE, VARCHAR )</param>
        /// <param name="parameterTyp">OleDb-Typ der Werte</param>
        /// <param name="zeilen">die Zeilen als Werte</param>
        /// <param name="kopf">Spaltennamen genau dieser Daten ( headersDouble bzw. headersString )</param>
        /// <returns>true = geschrieben, false = Dialog abgebrochen</returns>
        bool WriteTable( string newFileTarget, string newTableName, bool overwrite, string funktion, int columns, string typ,
            OleDbType parameterTyp, IEnumerable<object[]> zeilen, string[] kopf )
        {
            bool perDialog = ( newFileTarget == "" );
            // ohne Namen per Dialog; abgebrochen: nichts schreiben
            if ( !ChooseTarget( ref newFileTarget, !perDialog ) )
                return ( false );
            // vorhandene Datei: im Dialog hat Windows schon gefragt, sonst entscheidet 'overwrite';
            // geschrieben wird in eine Zwischendatei, die alte Datei wird erst danach ersetzt ( 28.11.26, Befund H )
            string zwischen = ExcelOleDb.PrepareTarget( newFileTarget, perDialog || overwrite, funktion );

            string commandCreate = ExcelOleDb.TableCommands( newTableName,
                ExcelOleDb.ColumnNamesForWrite( kopf, true, columns ), typ, out string commandInsert );
            try
            {   // Zwischendatei anlegen und füllen
                CreateDatabase( zwischen );
                using ( OleDbConnection connection = new OleDbConnection( targetConnectionStart + zwischen + targetConnectionOptions ) )
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand( commandCreate, connection );
                    ExcelOleDb.CreateTable( command, newTableName, funktion );
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
            }   // Ende: Zwischendatei anlegen und füllen
            catch
            {   // gescheitert: Zwischendatei weg, alte Zieldatei unberührt
                ExcelOleDb.DeleteQuiet( zwischen );
                throw;
            }   // Ende: gescheitert
            ExcelOleDb.CommitTarget( zwischen, newFileTarget, funktion );
            writtenTableName = newTableName;   // Access übernimmt den Namen unverändert ( gemessen 28.11.26 )
            return ( true );

        }   // Ende: bool WriteTable

        /// <summary>
        /// Intern data list double will be written into a new
        /// Access file. If not given a name a dialog will query for it ( 27.11.26: neue Datei per ADOX,
        /// vorher scheiterte jeder Aufruf mit "Datei nicht gefunden" ). Eine vorhandene Zieldatei wird nur ersetzt,
        /// wenn sie im Dialog bestätigt oder 'overwrite' gesetzt ist, und erst nach gelungenem Schreiben ( 28.11.26 ).
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Tabellenname: höchstens 64 Zeichen, ohne . ! ` [ ], kein führendes Leerzeichen</param>
        /// <param name="overwrite">true = eine vorhandene Zieldatei ohne Rückfrage ersetzen ( nur mit Namen von Bedeutung )</param>
        /// <returns>true = geschrieben; false = Dialog abgebrochen, nichts geschrieben</returns>
        /// <exception cref="InvalidOperationException">keine Zahlen zum Schreiben, oder die Tabelle ließ sich nicht anlegen</exception>
        /// <exception cref="IOException">Zieldatei vorhanden und nicht freigegeben, oder sie lässt sich nicht ersetzen; die alte bleibt dann unverändert</exception>
        public bool WriteListDoubleToNewTarget( string newFileTarget = "", string newTableName = "newDoubles", bool overwrite = false )
        {
            const string funktion = "AccessOleDb.WriteListDoubleToNewTarget";
            // die Daten haben zu passen: vor Dialog und Datei prüfen ( 28.11.26: Ausnahme statt MessageBox )
            writtenTableName = "";   // ganz oder gar nicht: erst nach gelungenem Schreiben gesetzt ( 28.11.26 )
            ExcelOleDb.CheckHasData( valuesDouble, funktion, "ReadDoubleList" );
            ExcelOleDb.CheckTableName( newTableName, ExcelOleDb.forbiddenTableCharsAccess, funktion );
            List<object[]> zeilen = new List<object[]>( );
            foreach ( double[] zeile in valuesDouble )
                zeilen.Add( zeile.Cast<object>( ).ToArray( ) );
            return ( WriteTable( newFileTarget, newTableName, overwrite, funktion, valuesDouble[ 0 ].Length, "DOUBLE",
                OleDbType.Double, zeilen, headersDouble ) );

        }   // end: WriteListDoubleToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new
        /// Access file. If not given a name a dialog will query for it. Texte höchstens 255 Zeichen
        /// ( VARCHAR, wie ExcelOleDb ), vorher geprüft, sonst nichts geschrieben. Eine vorhandene Zieldatei wird nur
        /// ersetzt, wenn sie im Dialog bestätigt oder 'overwrite' gesetzt ist, und erst nach gelungenem Schreiben ( 28.11.26 ).
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Tabellenname: höchstens 64 Zeichen, ohne . ! ` [ ], kein führendes Leerzeichen</param>
        /// <param name="overwrite">true = eine vorhandene Zieldatei ohne Rückfrage ersetzen ( nur mit Namen von Bedeutung )</param>
        /// <returns>true = geschrieben; false = Dialog abgebrochen, nichts geschrieben</returns>
        /// <exception cref="InvalidOperationException">keine Texte zum Schreiben, oder die Tabelle ließ sich nicht anlegen</exception>
        /// <exception cref="ArgumentException">ein Text ist länger als 'ExcelOleDb.maxTextLength', oder der Tabellenname ist ungültig</exception>
        /// <exception cref="IOException">Zieldatei vorhanden und nicht freigegeben, oder sie lässt sich nicht ersetzen; die alte bleibt dann unverändert</exception>
        public bool WriteListStringToNewTarget( string newFileTarget = "", string newTableName = "newStrings", bool overwrite = false )
        {
            const string funktion = "AccessOleDb.WriteListStringToNewTarget";
            // die Daten haben zu passen: vor Dialog und Datei prüfen ( wie ExcelOleDb; 28.11.26: Ausnahme statt MessageBox )
            writtenTableName = "";   // ganz oder gar nicht: erst nach gelungenem Schreiben gesetzt ( 28.11.26 )
            ExcelOleDb.CheckHasData( valuesString, funktion, "ReadStringList" );
            ExcelOleDb.CheckTableName( newTableName, ExcelOleDb.forbiddenTableCharsAccess, funktion );
            ExcelOleDb.CheckTextLength( valuesString, funktion );
            List<object[]> zeilen = new List<object[]>( );
            foreach ( string[] zeile in valuesString )
                zeilen.Add( zeile.Cast<object>( ).ToArray( ) );
            // Spalten der Texte ( 27.11.26: vorher aus valuesDouble, falsch bzw. Absturz ohne gelesene Zahlen )
            return ( WriteTable( newFileTarget, newTableName, overwrite, funktion, valuesString[ 0 ].Length, "VARCHAR",
                OleDbType.VarChar, zeilen, headersString ) );

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
#if !NETSTANDARD2_0
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
#else
            // netstandard2.0 ( 28.11.26 ): kein Dateidialog, false wie "abgebrochen"; Dateinamen direkt übergeben
#endif
            return ( false );

        }   // end: DialogFileNameLoad

        /// <summary>
        /// Queries a filename from the user with the standard dialog.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool DialogFileNameSave( ref string fileName )
        {
#if !NETSTANDARD2_0
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.OverwritePrompt = true;  // Windows fragt vor dem Ersetzen ( 28.11.26: einzige Rückfrage, Punkt l )
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
#else
            // netstandard2.0 ( 28.11.26 ): kein Dateidialog, false wie "abgebrochen"; Dateinamen direkt übergeben
#endif
            return ( false );

        }   // end: DialogFileNameLoad



    }   // end: public class AccessOleDb

}   // end: namespace DbaseFrame

