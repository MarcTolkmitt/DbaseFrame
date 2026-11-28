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
using System.IO;
using System.Net;
using System.Configuration;
using System.Data;

namespace DbaseFrame
{
    /// <summary>
    /// Reads and writes Excel files ( .xlsx ) through the ACE OleDb driver: table names, column names, texts,
    /// numbers and cell types of a table; writes one new table into a new file. History: see the field 'version'.
    /// </summary>
    public class ExcelOleDb
    {
        /// <summary>
        /// created on: 22.10.24
        /// last edit: 28.11.26
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
        ///           Read…List leeren ihre Liste zu Beginn ( nach einem Fehler keine Daten eines früheren Lesens );
        ///           ReadDoubleList liest mit IMEX=1 und wandelt je Zelle wie ExcelInterop ( Text statt Zahl →
        ///           InvalidDataException; vorher still 0, auch WAHR und "1,5" );
        ///           ReadStringList / ReadDoubleList: Tabelle ohne Datenzeile → InvalidDataException;
        ///           Write…ToNewTarget ohne Schließen/Öffnen zwischen CREATE und INSERT ( sonst Zahlen als Text )
        /// 28.11.26: ReadColumnNames( numTable ) neu ( wie AccessOleDb und ExcelInterop ); ReadTableNames, GetTableName
        ///           und ReadColumnNames zählen nur noch Arbeitsblätter ( SheetTables ), keine benannten Bereiche
        ///           ( Druckbereich, _xlnm#_FilterDatabase ); Reihenfolge des Treibers ( alphabetisch, nicht wie in Excel );
        ///           ReadColumnNames meldet ein leeres Blatt ( InvalidDataException, wie ExcelInterop );
        ///           Grenzen des ACE-Treibers dokumentiert ( gemessen, Entscheidung: dokumentieren statt umgehen ):
        ///           Text nach den ersten 8 Zeilen einer Zahlenspalte kommt als leerer Text ( ReadStringList ),
        ///           ein beschädigtes Blatt-XML wird ohne Meldung nur bis zur Bruchstelle gelesen;
        ///           Punkt l: keine MessageBox mehr aus der Bibliothek: keine Daten → InvalidOperationException
        ///           ( CheckHasData ), vorhandene Zieldatei nur nach Bestätigung im Dialog oder mit overwrite: true
        ///           ( PrepareTarget ), sonst IOException; Rückgabe false heißt nur noch „Dialog abgebrochen“;
        ///           Befund H: geschrieben wird in eine Zwischendatei, die alte Datei erst danach ersetzt ( CommitTarget ),
        ///           bei einem Fehler bleibt sie unverändert ( vorher vor dem Schreiben gelöscht );
        ///           Rückblick: Ordner ohne Schreibrecht → verständliche IOException ( PrepareTarget legt probeweise an );
        ///           Tabellenname vor dem Dialog geprüft ( CheckTableName, Regeln gemessen ), CREATE TABLE übersetzt;
        ///           neues Feld writtenTableName: Name der geschriebenen Tabelle, wie GetTableName / GetSheetName ihn liefert
        ///           28.11.26 ( 2 ): GetTableName wirft bei ungültiger Nummer ArgumentOutOfRangeException ( vorher leer ), neu GetTableCount
        /// </summary>
        Version version = new Version( "1.0.15" );

        // Connect to the Excel file
        /// <summary>
        /// First part of the connection string: provider and "Data Source=", followed by the file name.
        /// </summary>
        public string conStringStart =
            "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=";
        /// <summary>
        /// Part of the connection string after the file name: opens the Extended Properties for Excel 12.0 Xml.
        /// </summary>
        public string conStringEnd =
            ";Extended Properties=\"Excel 12.0 Xml;";
        /// <summary>
        /// Last part of the connection string when the first row is the header ( HDR=YES ).
        /// </summary>
        public string withHeader = "HDR=YES;\"";
        /// <summary>
        /// Last part of the connection string when there is no header row ( HDR=NO ).
        /// </summary>
        public string withoutHeader = "HDR=NO;\"";
        /// <summary>
        /// true = the first row of a table holds the column names ( set by the constructor ).
        /// </summary>
        public bool useHeader = true;
        /// <summary>
        /// Connection string of the source file: conStringStart + file + conStringEnd + withHeader or withoutHeader.
        /// </summary>
        public string connectionString = "";
        /// <summary>
        /// Connection string of the target file chosen by ChooseTarget ( always with header ).
        /// </summary>
        public string targetConnectionString = "";
        /// <summary>
        /// Name of the source file as given to the constructor or chosen in the dialog.
        /// </summary>
        public string fileName = "";

        /// <summary>
        /// Grund, wenn keine Datei gewählt ist ( z. B. "Dateiauswahl abgebrochen" ); sonst leer ( wie ExcelInterop ).
        /// </summary>
        public string lastError = "";
        /// <summary>
        /// Name of the target file chosen by ChooseTarget ( called by Write…ToNewTarget ).
        /// </summary>
        public string targetFileName = "";

        /// <summary>
        /// Name der zuletzt geschriebenen Tabelle, so wie GetTableName / ReadTableNames ihn liefern ( 28.11.26 ). Der
        /// ACE-Treiber legt das Blatt oft verändert an ( "a b" wird "a_b$", "1abc" wird "_1abc$", gemessen ), daher aus der
        /// geschriebenen Datei gelesen. Zu Beginn jedes Schreibens leer, nur nach gelungenem Schreiben gesetzt.
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
        /// Spaltennamen der zuletzt gelesenen Tabelle ( wie ExcelInterop.headers ): bei useHeader die Kopfzeile,
        /// sonst die Namen des Treibers ( F1, F2, ... ). Nur nach fehlerfreiem Lesen gesetzt, sonst leer.
        /// Die Schreibfunktionen nehmen sie bei useHeader als Spaltennamen der neuen Tabelle ( 27.11.26 ).
        /// </summary>
        public string[] headers = new string[ 0 ];

        // Spaltennamen je Datensatz ( 27.11.26 ): gesetzt zusammen mit valuesString bzw. valuesDouble, damit die
        // Schreibfunktionen die Kopfzeile genau dieser Daten nehmen, auch wenn danach ein anderes Lesen scheitert
        string[] headersString = new string[ 0 ];
        string[] headersDouble = new string[ 0 ];
        /// <summary>
        /// Names of the worksheets ( tables ) of the source file, filled by ReadTableNames.
        /// </summary>
        public string[] sheets = new string[1];
        /// <summary>
        /// Index of the chosen table in 'sheets'; -1 = none chosen yet.
        /// </summary>
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
        /// <param name="doUseHeader">true = the first row of a table holds the column names ( HDR=YES )</param>
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
#if !NETSTANDARD2_0
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
#else
            // netstandard2.0 ( 28.11.26 ): kein Dateidialog, false wie "abgebrochen"; Dateinamen direkt übergeben
#endif
            return ( false );

        }   // end: DialogFileNameLoad

        /* Die Datei hat zu passen, sonst Fehlermeldung:
         *  - fehlt sie, legt der ACE-Treiber beim Öffnen stillschweigend eine neue, leere Mappe an
         *    → vorher prüfen, FileNotFoundException
         *  - ist sie beschädigt, kein Excel-Format, kennwortgeschützt oder woanders gesperrt, wirft Open( ) eine OleDbException
         *    ( "Die externe Tabelle hat nicht das erwartete Format." ) → InvalidDataException mit Dateinamen
        */
        /// <summary>
        /// Öffnet die Verbindung zur Quelldatei zum Lesen.
        /// </summary>
        /// <returns>die geöffnete Verbindung ( mit using verwenden )</returns>
        /// <param name="imex">mit IMEX=1: gemischte Spalten kommen als Text statt als DBNull ( für ReadDoubleList )</param>
        OleDbConnection OpenSource( bool imex = false )
        {
            // keine Datei gewählt ( Dialog abgebrochen ): klare Meldung statt "nicht gefunden: " ohne Namen
            if ( sourceFile == "" )
                throw new InvalidOperationException( "ExcelOleDb: keine Datei gewählt ( " + lastError + " )." );
            if ( !File.Exists( sourceFile ) )
                throw new FileNotFoundException( $"Excel-Datei nicht gefunden: {sourceFile}", sourceFile );
            CheckReadable( sourceFile );
            // IMEX=1 kommt in die Extended Properties ( vor das schließende Anführungszeichen )
            string verbindung = imex ? connectionString.TrimEnd( '"' ) + "IMEX=1;\"" : connectionString;
            OleDbConnection conn = new OleDbConnection( verbindung );
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

        /* Grund ( 27.11.26, gemessen ): Hält ein anderer Prozess die Datei ohne Lesefreigabe offen
         * ( FileShare.None ), stürzt der ACE-Treiber in OleDbConnection.Open( ) mit einer Zugriffsverletzung
         * ( 0xC0000005 ) ab; das beendet das ganze Programm und lässt sich nicht abfangen.
         * Eine normal in Excel geöffnete Datei besteht die Prüfung; dafür meldet sich der Treiber selbst
         * ( "bereits exklusiv geöffnet", siehe OpenSource ).
        */
        /// <summary>
        /// Prüft vor dem Öffnen mit dem ACE-Treiber, ob sich die Datei lesen lässt.
        /// </summary>
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
        /// <param name="funktion">aufrufende Klasse und Funktion für die Meldung ( z. B. "ExcelOleDb.WriteList…" )</param>
        /// <exception cref="ArgumentException">ein Text ist zu lang, mit Position ( spalte, zeile ) ab 0</exception>
        internal static void CheckTextLength( List<string[]> rows, string funktion )
        {
            for ( int y = 0; y < rows.Count; y++ )
                for ( int x = 0; x < rows[ y ].Length; x++ )
                    if ( ( rows[ y ][ x ]?.Length ?? 0 ) > maxTextLength )
                        throw new ArgumentException( funktion + ": Text an ( " + x + ", " + y
                            + " ) hat " + rows[ y ][ x ].Length + " Zeichen, höchstens " + maxTextLength
                            + " sind möglich. Nichts geschrieben." );

        }   // Ende: void CheckTextLength

        /// <summary>
        /// Prüft vor dem Schreiben, dass es Daten gibt ( 28.11.26, Punkt l: ersetzt die MessageBox "No data to write"
        /// in ExcelOleDb und AccessOleDb; ExcelInterop meldet das schon so ).
        /// </summary>
        /// <typeparam name="T">double oder string</typeparam>
        /// <param name="rows">die Zeilen ( valuesDouble bzw. valuesString )</param>
        /// <param name="funktion">aufrufende Klasse und Funktion für die Meldung</param>
        /// <param name="lesen">Lesefunktion, die die Daten liefert ( für den Hinweis in der Meldung )</param>
        /// <exception cref="InvalidOperationException">keine Zeilen oder leere erste Zeile, nichts geschrieben</exception>
        internal static void CheckHasData<T>( List<T[]> rows, string funktion, string lesen )
        {
            if ( ( rows == null ) || ( rows.Count < 1 ) || ( rows[ 0 ].Length == 0 ) )
                throw new InvalidOperationException( funktion + ": keine Daten zum Schreiben ( zuerst " + lesen
                    + " fehlerfrei ausführen ). Nichts geschrieben." );

        }   // Ende: void CheckHasData

        /// <summary>
        /// Prüft die Zieldatei vor dem Schreiben und liefert die Zwischendatei, in die geschrieben wird ( 28.11.26,
        /// Punkt l und Befund H: ersetzt die Rückfrage per MessageBox und das Löschen vor dem Schreiben in allen drei
        /// Klassen ). Eine vorhandene Datei darf nur ersetzt werden, wenn 'ersetzen' gilt: im Speichern-Dialog bestätigt
        /// ( Windows fragt dort selbst, OverwritePrompt ) oder vom Aufrufer mit overwrite: true freigegeben. Sie wird hier
        /// nicht angerührt, nur geprüft, ob sie sich ersetzen lässt; ersetzt wird sie erst nach gelungenem Schreiben
        /// ( CommitTarget ). So bleibt bei jedem Fehler beim Schreiben die alte Datei erhalten.
        /// </summary>
        /// <param name="datei">Name der Zieldatei</param>
        /// <param name="ersetzen">true = eine vorhandene Datei darf ersetzt werden</param>
        /// <param name="funktion">aufrufende Klasse und Funktion für die Meldung</param>
        /// <returns>Name der Zwischendatei im selben Ordner, mit derselben Endung ( das Format hängt an der Endung )</returns>
        /// <exception cref="IOException">Datei vorhanden und nicht freigegeben, oder sie lässt sich nicht ersetzen
        /// ( geöffnet, schreibgeschützt )</exception>
        internal static string PrepareTarget( string datei, bool ersetzen, string funktion )
        {
            if ( File.Exists( datei ) )
            {   // vorhandene Zieldatei
                if ( !ersetzen )
                    throw new IOException( funktion + ": Die Zieldatei gibt es schon: " + datei
                        + " ( zum Ersetzen overwrite: true übergeben oder einen anderen Namen wählen ). Nichts geschrieben." );
                try
                {   // ersetzbar? exklusiv zum Schreiben öffnen und gleich schließen, Inhalt bleibt
                    using ( FileStream probe = new FileStream( datei, FileMode.Open, FileAccess.ReadWrite, FileShare.None ) )
                    {   // nur öffnen
                    }   // Ende: nur öffnen
                }   // Ende: ersetzbar?
                catch ( Exception fehler ) when ( ( fehler is IOException ) || ( fehler is UnauthorizedAccessException ) )
                {   // gesperrt oder schreibgeschützt: verständlich melden
                    throw new IOException( funktion + ": Die vorhandene Zieldatei lässt sich nicht ersetzen: " + datei
                        + " ( in einem anderen Programm geöffnet oder schreibgeschützt? ). Nichts geschrieben.", fehler );
                }   // Ende: gesperrt oder schreibgeschützt
            }   // Ende: vorhandene Zieldatei

            string ordner = Path.GetDirectoryName( datei ) ?? "";
            string name = Path.GetFileNameWithoutExtension( datei ) + ".~dbf" + Guid.NewGuid( ).ToString( "N" ).Substring( 0, 8 )
                + Path.GetExtension( datei );
            string zwischen = Path.Combine( ordner, name );
            try
            {   // darf im Ordner angelegt werden? Probe unter dem Namen der Zwischendatei, gleich wieder entfernt
                // ( 28.11.26 gemessen: sonst meldet der ACE-Treiber nur "kann die Datei '' nicht öffnen" )
                using ( FileStream probe = new FileStream( zwischen, FileMode.CreateNew, FileAccess.Write ) )
                {   // nur anlegen
                }   // Ende: nur anlegen
                File.Delete( zwischen );
            }   // Ende: darf im Ordner angelegt werden?
            catch ( Exception fehler ) when ( ( fehler is IOException ) || ( fehler is UnauthorizedAccessException ) )
            {   // Ordner fehlt oder keine Schreibrechte
                DeleteQuiet( zwischen );
                throw new IOException( funktion + ": Im Ordner '" + Path.GetFullPath( ordner == "" ? "." : ordner )
                    + "' lässt sich keine Datei anlegen ( Ordner fehlt oder keine Schreibrechte? ). Nichts geschrieben.", fehler );
            }   // Ende: Ordner fehlt oder keine Schreibrechte
            return ( zwischen );

        }   // Ende: string PrepareTarget

        /// <summary>
        /// Zeichen, die der ACE-Treiber in einem neuen Excel-Tabellennamen ablehnt ( gemessen 28.11.26 ).
        /// </summary>
        internal const string forbiddenTableCharsExcel = "!`[]:$";

        /// <summary>
        /// Zeichen, die der ACE-Treiber in einem neuen Access-Tabellennamen ablehnt ( gemessen 28.11.26 ).
        /// </summary>
        internal const string forbiddenTableCharsAccess = ".!`[]";

        /// <summary>
        /// Prüft den Namen einer neuen Tabelle vor Dialog und Datei ( 28.11.26, gemessen mit beiden ACE-Treibern; vorher
        /// kam nur "Syntaxfehler in CREATE TABLE" oder "ist kein gültiger Name" ): nicht leer, kein führendes
        /// Leerzeichen, höchstens 64 Zeichen, keines der verbotenen Zeichen.
        /// </summary>
        /// <param name="name">Tabellenname</param>
        /// <param name="verboten">forbiddenTableCharsExcel bzw. forbiddenTableCharsAccess</param>
        /// <param name="funktion">aufrufende Klasse und Funktion für die Meldung</param>
        /// <exception cref="ArgumentException">Name ungültig, mit Grund</exception>
        internal static void CheckTableName( string name, string verboten, string funktion )
        {
            string grund = "";
            if ( string.IsNullOrWhiteSpace( name ) )
                grund = "leer";
            else if ( name[ 0 ] == ' ' )
                grund = "beginnt mit einem Leerzeichen";
            else if ( name.Length > 64 )
                grund = name.Length + " Zeichen, höchstens 64";
            else if ( name.IndexOfAny( verboten.ToCharArray( ) ) >= 0 )
                grund = "enthält eines der Zeichen " + string.Join( " ", verboten.ToCharArray( ) );
            if ( grund != "" )
                throw new ArgumentException( funktion + ": Der Tabellenname '" + name + "' ist ungültig ( " + grund
                    + " ). Nichts geschrieben.", "newTableName" );

        }   // Ende: void CheckTableName

        /// <summary>
        /// Führt CREATE TABLE aus und übersetzt einen Fehler des Treibers ( 28.11.26, Rückfallebene zu CheckTableName ).
        /// </summary>
        /// <param name="command">Befehl mit CREATE TABLE, Verbindung offen</param>
        /// <param name="tabelle">Tabellenname für die Meldung</param>
        /// <param name="funktion">aufrufende Klasse und Funktion für die Meldung</param>
        /// <exception cref="InvalidOperationException">Tabelle ließ sich nicht anlegen</exception>
        internal static void CreateTable( OleDbCommand command, string tabelle, string funktion )
        {
            try
            {   // anlegen
                command.ExecuteNonQuery( );
            }   // Ende: anlegen
            catch ( OleDbException fehler )
            {   // verständlich melden
                throw new InvalidOperationException( funktion + ": Die Tabelle '" + tabelle + "' ließ sich nicht anlegen ( Treiber: "
                    + fehler.Message + " ). Nichts geschrieben.", fehler );
            }   // Ende: verständlich melden

        }   // Ende: void CreateTable

        /// <summary>
        /// Ersetzt nach gelungenem Schreiben die Zieldatei durch die Zwischendatei ( 28.11.26, Befund H ).
        /// Scheitert das, wird die Zwischendatei entfernt und die alte Zieldatei bleibt unverändert.
        /// </summary>
        /// <param name="zwischen">die fertig geschriebene Zwischendatei ( PrepareTarget )</param>
        /// <param name="datei">Name der Zieldatei</param>
        /// <param name="funktion">aufrufende Klasse und Funktion für die Meldung</param>
        /// <exception cref="IOException">Zieldatei ließ sich nicht ersetzen</exception>
        internal static void CommitTarget( string zwischen, string datei, string funktion )
        {
            try
            {   // ersetzen bzw. umbenennen
#if NETSTANDARD2_0
                // netstandard2.0 ( 28.11.26 ) kennt File.Move( a, b, true ) nicht: vorhandenes Ziel per File.Replace
                if ( File.Exists( datei ) )
                    File.Replace( zwischen, datei, null );
                else
                    File.Move( zwischen, datei );
#else
                File.Move( zwischen, datei, true );
#endif
            }   // Ende: ersetzen bzw. umbenennen
            catch ( Exception fehler ) when ( ( fehler is IOException ) || ( fehler is UnauthorizedAccessException ) )
            {   // aufräumen, alte Datei bleibt
                DeleteQuiet( zwischen );
                throw new IOException( funktion + ": Die Zieldatei ließ sich nach dem Schreiben nicht ersetzen: " + datei
                    + " ( inzwischen in einem anderen Programm geöffnet? ). Sie ist unverändert.", fehler );
            }   // Ende: aufräumen, alte Datei bleibt

        }   // Ende: void CommitTarget

        /// <summary>
        /// Entfernt eine Zwischendatei nach einem Fehler; ein Fehler dabei überdeckt nicht den eigentlichen.
        /// </summary>
        /// <param name="datei">die Zwischendatei</param>
        internal static void DeleteQuiet( string datei )
        {
            try
            {   // entfernen, falls vorhanden
                if ( File.Exists( datei ) )
                    File.Delete( datei );
            }   // Ende: entfernen, falls vorhanden
            catch ( Exception )
            {   // bleibt liegen; die eigentliche Meldung ist wichtiger
            }   // Ende: bleibt liegen

        }   // Ende: void DeleteQuiet

        // --------------------------------------------     the routines

        /// <summary>
        /// Reads the table's data types as anonymous array of
        /// strings into 'valuesTypes'.
        /// Grenze des ACE-Treibers wie bei ReadStringList: Text nach den ersten 8 Zeilen einer Zahlenspalte ergibt
        /// System.DBNull ( gemessen 28.11.26 ).
        /// </summary>
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
        /// Grenze des ACE-Treibers ( gemessen 28.11.26, gewollt dokumentiert statt umgangen ): er schätzt den Typ einer
        /// Spalte aus den ersten 8 Zeilen; steht danach Text in einer Zahlenspalte ( "keine Zahl", WAHR, "1,5" ), kommt er
        /// als DBNull, also als leerer Text, ohne Meldung. ExcelInterop nimmt jeden Zellwert selbst ( Value2 ) und
        /// schätzt keine Spaltentypen, passt also besser zu gemischten Spalten.
        /// </summary>
        public void ReadStringList( )
        {
            // ganz oder gar nicht ( 27.11.26, wie ExcelInterop ): zu Beginn leeren, lokal füllen, am Ende zuweisen
            headers = new string[ 0 ];
            valuesString = new List<string[]>();
            headersString = new string[ 0 ];
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
                // die Daten haben zu passen: eine Tabelle ohne Datenzeile ist kein Datensatz ( 27.11.26, wie ExcelInterop )
                if ( locStrings.Count == 0 )
                    throw new InvalidDataException( $"ExcelOleDb.ReadStringList: Tabelle '{SelectedTable( )}' enthält keine Daten"
                        + ( useHeader ? " unter der Kopfzeile." : "." ) );
                valuesString = locStrings;
                headersString = locHeaders;
                headers = locHeaders;

            }   // end: using

        }   // end: ReadStringList

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// doubles into 'valuesDouble'.
        /// </summary>
        public void ReadDoubleList( )
        {
            // ganz oder gar nicht ( 27.11.26, wie ExcelInterop ): zu Beginn leeren, lokal füllen, am Ende zuweisen
            headers = new string[ 0 ];
            valuesDouble = new List<double[]>();
            headersDouble = new string[ 0 ];
            /* Die Daten haben zu passen ( 27.11.26, wie ExcelInterop.ReadDoubles ):
             *  - ohne IMEX liefert der ACE-Treiber Text in einer Zahlenspalte als DBNull, nicht unterscheidbar von
             *    einer leeren Zelle ( gemessen ); mit IMEX=1 kommt eine gemischte Spalte als Text
             *  - je Zelle wie ExcelInterop: leer 0, WAHR/FALSCH 1/0, Zahl als Text umgewandelt, anderer Text Fehler
             *  - Grenze des Treibers: er schätzt den Spaltentyp aus den ersten 8 Zeilen ( TypeGuessRows ); steht
             *    Text erst danach in einer Zahlenspalte, kommt er weiter als DBNull, also 0. Streng: ExcelInterop
            */
            using ( OleDbConnection conn = OpenSource( true ) )
            {
                OleDbCommand command = new OleDbCommand($"SELECT * FROM [{SelectedTable( )}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                string[] locHeaders = ColumnNames( reader );

                List<double[]> locDoubles = new List<double[]>();

                while ( reader.Read() )
                {
                    double[] temp = new double[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                    {   // jeder Wert einzeln geprüft
                        if ( !TryCellNumber( reader[ pos ], out temp[ pos ] ) )
                            throw new InvalidDataException( $"ExcelOleDb.ReadDoubleList: Zeile {locDoubles.Count + ( useHeader ? 2 : 1 )}"
                                + $" der Tabelle, Spalte '{locHeaders[ pos ]}' enthält den Text '{reader[ pos ]}' statt einer Zahl"
                                + " ( nur leere Zellen werden 0 )." );
                    }   // Ende: jeder Wert einzeln geprüft
                    locDoubles.Add( temp );

                }
                // die Daten haben zu passen: eine Tabelle ohne Datenzeile ist kein Datensatz ( 27.11.26, wie ExcelInterop )
                if ( locDoubles.Count == 0 )
                    throw new InvalidDataException( $"ExcelOleDb.ReadDoubleList: Tabelle '{SelectedTable( )}' enthält keine Daten"
                        + ( useHeader ? " unter der Kopfzeile." : "." ) );
                valuesDouble = locDoubles;
                headersDouble = locHeaders;
                headers = locHeaders;

            }   // end: using

        }   // end: ReadDoubleList

        /// <summary>
        /// Zellwert als Zahl, Regeln wie ExcelInterop.TryCellNumber: Zahl unverändert, Datum als Seriennummer,
        /// WAHR/FALSCH 1/0 ( als Wahrheitswert oder Text WAHR, FALSCH, TRUE, FALSE ), Text mit einer Zahl in der
        /// Schreibweise von Windows ( "1,5" ) umgewandelt, leer ( DBNull, nur Leerzeichen ) 0; anderer Text: false.
        /// </summary>
        /// <param name="wert">Wert aus dem Leser</param>
        /// <param name="zahl">die Zahl</param>
        /// <returns>false bei Text, der keine Zahl ist</returns>
        internal static bool TryCellNumber( object? wert, out double zahl )
        {
            zahl = 0.0;
            switch ( wert )
            {   // je nach Typ des Werts
                case null:
                case DBNull:
                    return ( true );       // leer: Standardwert 0
                case double wertZahl:
                    zahl = wertZahl;
                    return ( true );
                case bool wahr:
                    zahl = wahr ? 1.0 : 0.0;
                    return ( true );
                case DateTime datum:
                    zahl = datum.ToOADate( );
                    return ( true );
                case string text:
                    string t = text.Trim( );
                    if ( t.Length == 0 )
                        return ( true );   // leer: Standardwert 0
                    if ( t.Equals( "WAHR", StringComparison.OrdinalIgnoreCase ) || t.Equals( "TRUE", StringComparison.OrdinalIgnoreCase ) )
                    {   // Wahrheitswert als Text ( IMEX liefert ihn so )
                        zahl = 1.0;
                        return ( true );
                    }   // Ende: Wahrheitswert als Text
                    if ( t.Equals( "FALSCH", StringComparison.OrdinalIgnoreCase ) || t.Equals( "FALSE", StringComparison.OrdinalIgnoreCase ) )
                        return ( true );   // 0
                    return ( double.TryParse( t, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.CurrentCulture, out zahl ) );
                case IConvertible sonst:
                    zahl = sonst.ToDouble( System.Globalization.CultureInfo.CurrentCulture );   // decimal, int, ...
                    return ( true );
                default:
                    return ( false );
            }   // Ende: je nach Typ des Werts

        }   // Ende: bool TryCellNumber

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
        /// Auch von AccessOleDb benutzt ( 27.11.26 ).
        /// </summary>
        /// <param name="kopfzeile">gelesene Spaltennamen ( 'headers' )</param>
        /// <param name="nutzeKopf">Spaltennamen übernehmen ( ExcelOleDb: useHeader )</param>
        /// <param name="columns">Spaltenzahl der Daten</param>
        /// <returns>je Spalte ein eindeutiger Name</returns>
        internal static string[] ColumnNamesForWrite( string[] kopfzeile, bool nutzeKopf, int columns )
        {
            bool kopf = nutzeKopf && ( kopfzeile.Length == columns );
            string[] headers = kopfzeile;
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
        /// <param name="namen">Spaltennamen ( ColumnNamesForWrite )</param>
        /// <param name="typ">SQL-Typ der Spalten ( DOUBLE, VARCHAR )</param>
        /// <param name="commandInsert">INSERT INTO mit Parametern @0, @1, ...</param>
        /// <returns>CREATE TABLE</returns>
        internal static string TableCommands( string tabelle, string[] namen, string typ, out string commandInsert )
        {
            int columns = namen.Length;
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
            // nur Arbeitsblätter ( SheetTables, 28.11.26 ), Zählung gleich in GetTableName und ReadColumnNames
            using ( OleDbConnection conn = OpenSource( ) )
            {
                sheets = SheetTables( conn );
            }   // end: using

#if NETSTANDARD2_0
            sheetNumber = -1;   // netstandard2.0 ( 28.11.26 ): kein Auswahldialog; Tabelle über sheetNumber bzw. SelectSheet wählen
#else
            DialogTablesChoice choice = new DialogTablesChoice( sheets );
            sheetNumber = choice.index;
#endif
            return( sheetNumber );

        }   // end: ReadTableNames

        /// <summary>
        /// Direct query for the table name ( worksheets only, alphabetical like the driver delivers them ).
        /// </summary>
        /// <param name="numTable">number of the sheet, 0 = first; valid up to GetTableCount( ) - 1</param>
        /// <returns>the name, e.g. "table0$"</returns>
        /// <exception cref="ArgumentOutOfRangeException">no table with this number ( since 1.0.9; before: empty text )</exception>
        public string GetTableName( int numTable )
        {
            string[] tabellen;
            using ( OleDbConnection conn = OpenSource( ) )
            {
                tabellen = SheetTables( conn );
            }   // end: using

            // 28.11.26: ungültige Nummer → Ausnahme wie ReadColumnNames ( Entscheidung „Angleichen“; vorher leerer Text )
            if ( ( numTable < 0 ) || ( numTable >= tabellen.Length ) )
                throw new ArgumentOutOfRangeException( nameof( numTable ), numTable,
                    $"ExcelOleDb.GetTableName: keine Tabelle Nr. {numTable} ( vorhanden: {tabellen.Length}, gezählt ab 0 )." );
            return ( tabellen[ numTable ] );

        }   // end: GetTableName

        /// <summary>
        /// Number of tables ( worksheets ) in the source file, counted like GetTableName / ReadColumnNames.
        /// </summary>
        /// <returns>the number of tables</returns>
        public int GetTableCount( )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                return ( SheetTables( conn ).Length );
            }   // end: using

        }   // end: int GetTableCount

        /// <summary>
        /// Die Arbeitsblätter der Mappe, wie der ACE-Treiber sie als Tabellen kennt ( "Tabelle1$", "'Mein Blatt$'" ),
        /// ohne benannte Bereiche ( "Druckbereich", "Tabelle1$_xlnm#_FilterDatabase" ). Reihenfolge des Treibers:
        /// alphabetisch, nicht die Blattreihenfolge in Excel ( anders als ExcelInterop, 28.11.26 ).
        /// </summary>
        /// <param name="conn">geöffnete Verbindung</param>
        /// <returns>Tabellennamen der Blätter</returns>
        static string[] SheetTables( OleDbConnection conn )
        {
            DataTable dt = conn.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null ) ?? new DataTable( );
            List<string> namen = new List<string>( );
            foreach ( DataRow zeile in dt.Rows )
            {   // Blätter enden auf $ ( bei Sonderzeichen in Hochkommas: '…$' )
                string name = zeile[ "TABLE_NAME" ]?.ToString( ) ?? string.Empty;
                if ( name.TrimEnd( '\'' ).EndsWith( "$" ) )
                    namen.Add( name );
            }   // Ende: Blätter enden auf $
            return ( namen.ToArray( ) );

        }   // Ende: string[] SheetTables

        /// <summary>
        /// Spaltennamen einer Tabelle, ohne Daten zu lesen ( 28.11.26, wie AccessOleDb und ExcelInterop ).
        /// Zählung der Blätter wie ReadTableNames / GetTableName ( SheetTables: alphabetisch ). Bei useHeader die
        /// Kopfzeile, sonst F1, F2, ...
        /// ( dieselben Namen, die ein Read…List in 'headers' liefert ).
        /// </summary>
        /// <param name="numTable">Nummer der Tabelle, 0 = erste</param>
        /// <returns>die Spaltennamen</returns>
        /// <exception cref="ArgumentOutOfRangeException">kein Blatt mit dieser Nummer</exception>
        /// <exception cref="InvalidDataException">leeres Blatt ( wie ExcelInterop )</exception>
        public string[] ReadColumnNames( int numTable )
        {
            using ( OleDbConnection conn = OpenSource( ) )
            {
                string[] tabellen = SheetTables( conn );
                if ( ( numTable < 0 ) || ( numTable >= tabellen.Length ) )
                    throw new ArgumentOutOfRangeException( nameof( numTable ),
                        $"ExcelOleDb: Blatt Nr. {numTable} gibt es nicht ( {tabellen.Length} Blätter in {sourceFile} )." );
                string tabelle = tabellen[ numTable ];

                string[] namen;
                using ( OleDbCommand command = new OleDbCommand( $"SELECT * FROM [{tabelle}]", conn ) )
                using ( OleDbDataReader reader = command.ExecuteReader( CommandBehavior.SchemaOnly ) )
                {   // nur das Schema, keine Datenzeile
                    namen = ColumnNames( reader );
                }   // Ende: nur das Schema, keine Datenzeile

                /* Leeres Blatt ( 28.11.26, wie ExcelInterop: die Daten haben zu passen ):
                 *  - der Treiber meldet dann eine einzige Spalte "F1" ohne Wert ( Annahme, mit Leer_Test.xlsx zu prüfen )
                 *  - nur in diesem Fall die Werte ansehen; eine Spalte mit Werten oder mit echter Kopfzeile ist kein leeres Blatt
                */
                if ( ( namen.Length == 1 ) && ( namen[ 0 ] == "F1" ) && !HasValue( conn, tabelle ) )
                    throw new InvalidDataException( $"ExcelOleDb: Blatt {tabelle} in {sourceFile} enthält keine Daten." );
                return ( namen );
            }   // Ende: using

        }   // Ende: string[] ReadColumnNames

        /// <summary>
        /// Enthält die erste Spalte der Tabelle irgendeinen Wert ( für die Erkennung eines leeren Blatts )?
        /// </summary>
        /// <param name="conn">geöffnete Verbindung</param>
        /// <param name="tabelle">Tabellenname</param>
        /// <returns>true, sobald ein Wert gefunden ist</returns>
        static bool HasValue( OleDbConnection conn, string tabelle )
        {
            using ( OleDbCommand command = new OleDbCommand( $"SELECT * FROM [{tabelle}]", conn ) )
            using ( OleDbDataReader reader = command.ExecuteReader( ) )
            {   // bis zum ersten Wert
                while ( reader.Read( ) )
                    if ( !reader.IsDBNull( 0 ) )
                        return ( true );
            }   // Ende: bis zum ersten Wert
            return ( false );

        }   // Ende: bool HasValue

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
        /// Eine vorhandene Zieldatei wird nur ersetzt, wenn sie im Dialog bestätigt oder 'overwrite' gesetzt ist
        /// ( 28.11.26, ohne MessageBox aus der Bibliothek ). Ersetzt wird erst nach gelungenem
        /// Schreiben ( über eine Zwischendatei ); bei einem Fehler bleibt die alte Datei unverändert.
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Name der neuen Tabelle: höchstens 64 Zeichen, ohne ! ` [ ] : $, kein führendes Leerzeichen.</param>
        /// <param name="overwrite">true = eine vorhandene Zieldatei ohne Rückfrage ersetzen ( nur mit Namen von Bedeutung )</param>
        /// <returns>true = geschrieben; false = Dialog abgebrochen, nichts geschrieben</returns>
        /// <exception cref="InvalidOperationException">keine Zahlen zum Schreiben, oder die Tabelle ließ sich nicht anlegen</exception>
        /// <exception cref="IOException">Zieldatei vorhanden und nicht freigegeben, oder sie lässt sich nicht ersetzen; die alte bleibt dann unverändert</exception>
        public bool WriteListDoubleToNewTarget( string newFileTarget = "", string newTableName = "newDoubles", bool overwrite = false )
        {
            // Ergebnis ganz oder gar nicht: Name erst nach gelungenem Schreiben ( 28.11.26 )
            writtenTableName = "";
            // die Daten haben zu passen: vor Dialog und Datei prüfen ( 28.11.26: Ausnahme statt MessageBox )
            CheckHasData( valuesDouble, "ExcelOleDb.WriteListDoubleToNewTarget", "ReadDoubleList" );
            CheckTableName( newTableName, forbiddenTableCharsExcel, "ExcelOleDb.WriteListDoubleToNewTarget" );

            bool perDialog = ( newFileTarget == "" );
            // ohne Namen per Dialog; abgebrochen: nichts schreiben
            if ( !ChooseTarget( ref newFileTarget, !perDialog ) )
                return ( false );
            // vorhandene Datei: im Dialog hat Windows schon gefragt, sonst entscheidet 'overwrite';
            // geschrieben wird in eine Zwischendatei, die alte Datei wird erst danach ersetzt ( 28.11.26, Befund H )
            const string funktion = "ExcelOleDb.WriteListDoubleToNewTarget";
            string zwischen = PrepareTarget( newFileTarget, perDialog || overwrite, funktion );

            // craft the 'CREATE TABLE' and 'INSERT INTO'; Spaltennamen bei useHeader aus der gelesenen
            // Kopfzeile ( 27.11.26: vorher immer 0, 1, 2, ... )
            int columns =  valuesDouble[0].Length;
            string commandCreate = TableCommands( newTableName, ColumnNamesForWrite( headersDouble, useHeader, columns ), "DOUBLE",
                out string commandInsert );
            string neuerName = string.Empty;
            try
            {   // in die Zwischendatei schreiben
                using ( OleDbConnection connection = new OleDbConnection( conStringStart + zwischen + conStringEnd + withHeader ) )
                {
                    connection.Open();
                    // create the table
                    OleDbCommand command = new OleDbCommand( commandCreate, connection );
                    CreateTable( command, newTableName, funktion );
                    // Verbindung offen lassen ( 27.11.26, gemessen ): Schließen und Wiederöffnen zwischen CREATE TABLE und
                    // INSERT ließ den ACE-Treiber die Zahlen als Text schreiben ( "1,5" statt 1.5 )
                    foreach ( double[] row in valuesDouble )
                    {
                        command.CommandText = commandInsert;
                        command.Parameters.Clear();

                        for ( int pos = 0; pos < row.Length; pos++ )
                            //command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );
                            command.Parameters.Add( $"@{pos}", OleDbType.Double ).Value = row[ pos ];

                        command.ExecuteNonQuery();
                    }

                    // tatsächlicher Blattname, wie der Treiber ihn angelegt hat ( 28.11.26 )
                    string[ ] angelegt = SheetTables( connection );
                    neuerName = ( angelegt.Length == 1 ) ? angelegt[ 0 ] : string.Empty;
                    connection.Close();

                }
            }   // Ende: in die Zwischendatei schreiben
            catch
            {   // gescheitert: Zwischendatei weg, alte Zieldatei unberührt
                DeleteQuiet( zwischen );
                throw;
            }   // Ende: gescheitert
            CommitTarget( zwischen, newFileTarget, funktion );
            writtenTableName = neuerName;
            return ( true );

        }   // end: WriteListDoubleToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new
        /// Excel file. If not given a name a dialog will query for it.
        /// Eine vorhandene Zieldatei wird nur ersetzt, wenn sie im Dialog bestätigt oder 'overwrite' gesetzt ist
        /// ( 28.11.26, ohne MessageBox aus der Bibliothek ). Ersetzt wird erst nach gelungenem
        /// Schreiben ( über eine Zwischendatei ); bei einem Fehler bleibt die alte Datei unverändert.
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Name der neuen Tabelle: höchstens 64 Zeichen, ohne ! ` [ ] : $, kein führendes Leerzeichen.</param>
        /// <param name="overwrite">true = eine vorhandene Zieldatei ohne Rückfrage ersetzen ( nur mit Namen von Bedeutung )</param>
        /// <returns>true = geschrieben; false = Dialog abgebrochen, nichts geschrieben</returns>
        /// <exception cref="InvalidOperationException">keine Texte zum Schreiben, oder die Tabelle ließ sich nicht anlegen</exception>
        /// <exception cref="ArgumentException">ein Text ist länger als 'maxTextLength', oder der Tabellenname ist ungültig</exception>
        /// <exception cref="IOException">Zieldatei vorhanden und nicht freigegeben, oder sie lässt sich nicht ersetzen; die alte bleibt dann unverändert</exception>
        public bool WriteListStringToNewTarget( string newFileTarget = "", string newTableName = "newStrings", bool overwrite = false )
        {
            // Ergebnis ganz oder gar nicht: Name erst nach gelungenem Schreiben ( 28.11.26 )
            writtenTableName = "";
            // die Daten haben zu passen: vor Dialog und Datei prüfen ( vorher Abbruch mitten im Schreiben,
            // die Zieldatei blieb mit leerer Tabelle zurück; 28.11.26: keine Daten → Ausnahme statt MessageBox )
            CheckHasData( valuesString, "ExcelOleDb.WriteListStringToNewTarget", "ReadStringList" );
            CheckTableName( newTableName, forbiddenTableCharsExcel, "ExcelOleDb.WriteListStringToNewTarget" );
            CheckTextLength( valuesString, "ExcelOleDb.WriteListStringToNewTarget" );

            bool perDialog = ( newFileTarget == "" );
            // ohne Namen per Dialog; abgebrochen: nichts schreiben
            if ( !ChooseTarget( ref newFileTarget, !perDialog ) )
                return ( false );
            // vorhandene Datei: im Dialog hat Windows schon gefragt, sonst entscheidet 'overwrite';
            // geschrieben wird in eine Zwischendatei, die alte Datei wird erst danach ersetzt ( 28.11.26, Befund H )
            const string funktion = "ExcelOleDb.WriteListStringToNewTarget";
            string zwischen = PrepareTarget( newFileTarget, perDialog || overwrite, funktion );

            // craft the 'CREATE TABLE' and 'INSERT INTO'
            // Spalten der Texte ( 27.11.26: vorher aus valuesDouble, falsch bzw. Absturz ohne gelesene Zahlen );
            // Spaltennamen bei useHeader aus der gelesenen Kopfzeile ( 27.11.26: vorher immer 0, 1, 2, ... )
            int columns =  valuesString[0].Length;
            string commandCreate = TableCommands( newTableName, ColumnNamesForWrite( headersString, useHeader, columns ), "VARCHAR",
                out string commandInsert );
            string neuerName = string.Empty;
            try
            {   // in die Zwischendatei schreiben
                using ( OleDbConnection connection = new OleDbConnection( conStringStart + zwischen + conStringEnd + withHeader ) )
                {
                    connection.Open();
                    // create the table
                    OleDbCommand command = new OleDbCommand( commandCreate, connection );
                    CreateTable( command, newTableName, funktion );
                    // Verbindung offen lassen wie in WriteListDoubleToNewTarget ( 27.11.26 )
                    foreach ( string[] row in valuesString )
                    {
                        command.CommandText = commandInsert;
                        command.Parameters.Clear();

                        for ( int pos = 0; pos < row.Length; pos++ )
                            //command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );
                            command.Parameters.Add( $"@{pos}", OleDbType.VarChar ).Value = row[ pos ];

                        command.ExecuteNonQuery();
                    }

                    // tatsächlicher Blattname, wie der Treiber ihn angelegt hat ( 28.11.26 )
                    string[ ] angelegt = SheetTables( connection );
                    neuerName = ( angelegt.Length == 1 ) ? angelegt[ 0 ] : string.Empty;
                    connection.Close();
                }
            }   // Ende: in die Zwischendatei schreiben
            catch
            {   // gescheitert: Zwischendatei weg, alte Zieldatei unberührt
                DeleteQuiet( zwischen );
                throw;
            }   // Ende: gescheitert
            CommitTarget( zwischen, newFileTarget, funktion );
            writtenTableName = neuerName;
            return ( true );

        }   // end: WriteListStringToNewTarget

    }   // end: ExcelOleDb

}   // end: namespace DbaseFrame
