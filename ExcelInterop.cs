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
using System.Configuration;
using System.Data;
using System.Globalization;
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
        /// last edit: 28.11.26
        /// Now since Visual Studio Pro i can use the Interop library
        /// 27.11.26: Lesefunktionen repariert ( ReadTypes, ReadStrings, ReadDoubles, je auch für einen Bereich:
        ///           Blatt, Value2, Index ab 0, X = Spalte / Y = Zeile, Bereichsprüfung ), Kopfzeile in 'headers'
        ///           wie ExcelOleDb; überall ActiveSheet.Worksheet ersetzt; Rückblick: Datenrechteck statt ab A1,
        ///           alle Werte mit einem Zugriff und danach einzeln geprüft, Sonderwerte ( Fehler, WAHR, Zahl als Text );
        ///           Schreibfunktionen repariert ( Blatt, Übergabe Matrix( x, y ) an Excels [ zeile, spalte ],
        ///           Anhängen unter dem Datenrechteck, WriteToNewTarget schreibt und speichert die neue Mappe );
        ///           die Daten haben zu passen: leeres Blatt / leere Rand-Zeile oder -Spalte und NaN / Unendlich
        ///           ergeben eine Fehlermeldung; Zielname wird vervollständigt ( .xlsx, voller Pfad );
        ///           Konstruktor mit doReadOnly ( schreibgeschützt öffnen, beim Beenden nicht speichern ),
        ///           workbookOpen; Blätter 0-basiert ( GetSheetName, SetSheetName, NewSheet ), SelectSheet,
        ///           Lesen und Schreiben im gewählten Blatt;
        ///           Grund eines gescheiterten Öffnens in 'lastError' ( fehlende / beschädigte Datei, Excel fehlt );
        ///           Öffnen ohne unsichtbare Rückfragen ( kennwortgeschützte Datei ließ das Programm hängen ),
        ///           relativer Dateiname → voller Pfad;
        ///           ExcelInteropQuit gibt auch das Blatt frei ( Excel lief nach dem Lesen weiter );
        ///           Write…ToNewTarget: Dialog abgebrochen → nichts schreiben, Rückgabe false ( vorher NewTarget.xlsx );
        ///           ReadDoubles: Text statt Zahl → InvalidDataException mit Zelladresse ( vorher still 0 );
        ///           ReadDoubles leert 'valuesDoubleArray' zu Beginn und weist es erst nach dem Lesen zu ( nach einem
        ///           Fehler keine halben oder alten Zahlen ); Schreiben ohne Zahlen → InvalidOperationException;
        ///           ReadStrings und ReadTypes ebenso, Schreiben ohne Texte → InvalidOperationException ( CheckStrings );
        ///           'headers' ebenso: zu Beginn leer, nur nach fehlerfreiem Lesen gesetzt ( ReadHeaders liefert zurück );
        ///           Write…ToNewTarget schreiben bei useHeader die Kopfzeile mit ( WriteHeaderRow, vorher fehlte sie );
        ///           ganz gesperrte Datei → lastError "gesperrt" ( vorher "beschädigt, kein Excel-Format" );
        ///           ReadStrings / ReadDoubles: Blatt nur mit Kopfzeile → Fehlermeldung ( CheckDataRows )
        /// 28.11.26: ReadColumnNames( numTable ) neu ( wie AccessOleDb und ExcelOleDb ), Blattwahl bleibt unverändert
        /// </summary>
        Version version = new Version( "1.0.9" );

        Microsoft.Office.Interop.Excel.Application excelApp = null!;
        Microsoft.Office.Interop.Excel.Workbook workbook = null!;
        Microsoft.Office.Interop.Excel.Worksheet worksheet = null!;

        public bool useHeader = true;
        public string fileName = "";
        public string targetFileName = "";
        public string[,] valuesStringArray = null!;
        public double[,] valuesDoubleArray = null!;
        public string[,] valuesTypesArray = null!;
        public string[] sheets = new string[1];
        public int sheetNumber = -1;

        /// <summary>
        /// Kopfzeile der zuletzt gelesenen Tabelle ( bei useHeader = false die Spaltennamen A, B, ... )
        /// </summary>
        public string[] headers = new string[ 0 ];

        // Kopfzeile je Datensatz ( 27.11.26 ): gesetzt zusammen mit valuesStringArray bzw. valuesDoubleArray, damit
        // Write…ToNewTarget die Kopfzeile genau dieser Daten schreiben, auch wenn danach ein anderes Lesen scheitert
        string[] headersString = new string[ 0 ];
        string[] headersDouble = new string[ 0 ];

        /// <summary>
        /// Mappe nur zum Lesen geöffnet: Excel öffnet sie schreibgeschützt, beim Beenden wird nichts gespeichert
        /// </summary>
        public bool readOnly = false;

        /// <summary>
        /// Mappe erfolgreich geöffnet?
        /// </summary>
        public bool workbookOpen = false;

        /// <summary>
        /// Grund, warum Excel nicht startete oder die Mappe nicht geöffnet wurde ( "" = kein Fehler );
        /// bei geöffneter Mappe ein Hinweis, z. B. dass sie gesperrt war und nur zum Lesen geöffnet wurde
        /// </summary>
        public string lastError = "";

        /// <summary>
        /// Constructor for the class.
        /// </summary>
        /// <param name="file">a file name</param>
        /// <param name="silent">query for the name via dialog ?</param>
        /// <param name="doUseHeader">erste Zeile des Datenrechtecks als Kopfzeile lesen</param>
        /// <param name="doReadOnly">nur lesen: schreibgeschützt öffnen, beim Beenden nichts speichern</param>
        public ExcelInterop( string file = "", bool silent = true,  bool doUseHeader = true, bool doReadOnly = false )
        {
            fileName = file;
            useHeader = doUseHeader;
            readOnly = doReadOnly;
            if ( !silent )
            {   // Dateiname per Dialog
                if ( !DialogFileNameLoad( ref fileName ) )
                {   // abgebrochen: nichts öffnen ( 27.11.26: vorher wurde trotzdem geöffnet )
                    lastError = "Dateiauswahl abgebrochen";
                    Console.WriteLine( "File dialog error... " );
                    return;
                }   // ende: abgebrochen: nichts öffnen
            }   // ende: Dateiname per Dialog
            try
            {
                excelApp =
                    new Microsoft.Office.Interop.Excel.Application();

            }
            catch ( COMException comEx )
            {
                lastError = "Excel ließ sich nicht starten: " + comEx.Message;
                Console.WriteLine( "Excel-Interop-error: " + comEx.Message );
            }
            catch ( Exception ex )
            {
                lastError = "Excel ließ sich nicht starten: " + ex.Message;
                Console.WriteLine( "common error: " + ex.Message );
            }   // end: try excelApp

            // Open document ( 27.11.26: über OpenWorkbook, bei readOnly schreibgeschützt; die Meldung
            // "File dialog error" kam vorher im stillen Modus immer, auch bei Erfolg )
            OpenWorkbook( );

        }   // end: ExcelInterop ( constructor )

        public void ExcelInteropQuit()
        {
            /* Alle festgehaltenen COM-Objekte freigeben, sonst läuft Excel nach dem Programmende weiter:
             *  - .NET führt beim Programmende keine Finalizer aus; das Feld 'worksheet' ( gesetzt beim Lesen und
             *    Schreiben ) hielt Excel am Leben, bis Windows den toten Verweis nach Minuten bemerkte
             *    ( 27.11.26 mit MatrixFFN gefunden: nach dem Lesen blieb je Aufruf ein Excel-Prozess )
             *  - Felder danach auf null: ein zweiter Aufruf tut nichts
             *  - GC zweimal: erst Finalizer der übrigen Zwischenobjekte ( UsedRange, Worksheets … ), dann deren Reste
            */
            if (worksheet != null)
            {   // gewähltes Blatt freigeben
                Marshal.ReleaseComObject(worksheet);
                worksheet = null!;
            }   // Ende: gewähltes Blatt freigeben

            // Optional: Workbook schließen
            if (workbook != null)
            {
                // nur gelesen: nichts speichern, die Datei bleibt unverändert ( 27.11.26: vorher immer gespeichert )
                workbook.Close(SaveChanges: !readOnly);
                Marshal.ReleaseComObject(workbook);
                workbook = null!;
            }
            workbookOpen = false;

            // Excel beenden
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null!;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
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
        /// <returns>false, wenn der Dialog abgebrochen wurde: dann nichts schreiben</returns>
        public bool ChooseTarget( ref string file, bool silent = true )
        {
            if ( !silent && !DialogFileNameSave( ref file ) )
            {   // Dialog abgebrochen ( 27.11.26: vorher wurde dann NewTarget.xlsx im Programmordner beschrieben )
                return ( false );
            }   // Ende: Dialog abgebrochen
            targetFileName = file;
            return ( true );

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
        /// Bei readOnly schreibgeschützt. Setzt 'workbookOpen'.
        /// </summary>
        /// <returns>boolean value indicating success or failure</returns>
        public bool OpenWorkbook()
        {
            workbookOpen = false;
            if ( excelApp == null )
                return ( false );   // Excel ließ sich nicht starten, Grund steht schon in lastError
            // relativer Name → voller Pfad: Excel löst relative Namen in seinem eigenen Ordner auf ( Dokumente ),
            // nicht im Arbeitsordner des Programms
            if ( fileName != "" )
                fileName = Path.GetFullPath( fileName );
            if ( !File.Exists( fileName ) )
            {   // fehlende Datei klar melden ( Excels Meldung dazu ist missverständlich )
                lastError = $"Excel-Datei nicht gefunden: {fileName}";
                Console.WriteLine( lastError );
                return ( false );
            }   // Ende: fehlende Datei klar melden
            try
            {   // ganz gesperrte Datei klar melden ( 27.11.26: Excel nannte sie "beschädigt, kein Excel-Format" )
                ExcelOleDb.CheckReadable( fileName );
            }
            catch ( IOException ioEx )
            {
                lastError = ioEx.Message;
                Console.WriteLine( lastError );
                return ( false );
            }   // Ende: ganz gesperrte Datei klar melden
            /* Keine unsichtbaren Rückfragen beim Öffnen, Excel läuft unsichtbar und würde ewig warten:
             *  - kennwortgeschützte Datei: ohne Kennwort fragt Excel danach → Platzhalter-Kennwort, dann meldet
             *    Excel einen Fehler ( bei ungeschützten Dateien wird das Kennwort nicht beachtet )
             *  - Schreibschutz-Kennwort ebenso ( beschreibbar öffnen scheitert dann, nur lesen geht ),
             *    Empfehlung "schreibgeschützt öffnen", Verknüpfungen aktualisieren
             *  - DisplayAlerts aus für alle übrigen Meldungen, danach wie vorher
             *  - das Platzhalter-Kennwort hat höchstens 15 Zeichen: mit 16 und mehr scheitert jede Datei
             *    ( mit Excel geprüft 27.11.26 )
            */
            const string keinKennwort = "DbF-kein-KW";
            bool alteRueckfragen = excelApp.DisplayAlerts;
            try
            {
                excelApp.DisplayAlerts = false;
                // Datei öffnen; bei readOnly schreibgeschützt ( 27.11.26: vorher immer beschreibbar )
                workbook = excelApp.Workbooks.Open(
                    fileName,
                    UpdateLinks: 0,
                    ReadOnly: readOnly,
                    Password: keinKennwort,
                    WriteResPassword: keinKennwort,
                    IgnoreReadOnlyRecommended: true,
                    Editable: !readOnly,
                    Notify: false
                );

            }
            catch ( COMException comEx )
            {   // z. B. beschädigt, kein Excel-Format, kennwortgeschützt; Excel meldet dann oft nur "0x800A03EC"
                lastError = "Excel-Datei kann nicht geöffnet werden ( beschädigt, kein Excel-Format, "
                    + ( readOnly ? "kennwortgeschützt" : "kennwort- oder schreibgeschützt" ) + " ): "
                    + $"{fileName}\n" + comEx.Message;
                Console.WriteLine( "Excel-Interop-Fehler: " + comEx.Message );
                return ( false );
            }   // Ende: z. B. beschädigt, kein Excel-Format, kennwortgeschützt
            catch ( Exception ex )
            {
                lastError = $"Excel-Datei kann nicht geöffnet werden: {fileName}\n" + ex.Message;
                Console.WriteLine( "Allgemeiner Fehler: " + ex.Message );
                return ( false );
            }   // end: try workbook
            finally
            {   // Rückfragen wie vorher
                excelApp.DisplayAlerts = alteRueckfragen;
            }   // Ende: Rückfragen wie vorher
            lastError = "";
            if ( !readOnly && workbook.ReadOnly )
            {   // woanders geöffnet ( gesperrt ): Excel öffnet dann ohne Rückfrage nur zum Lesen
                readOnly = true;
                lastError = $"Excel-Datei ist woanders geöffnet, nur zum Lesen geöffnet: {fileName}";
                Console.WriteLine( lastError );
            }   // Ende: woanders geöffnet ( gesperrt )
            workbookOpen = true;
            return ( true );

        }   // end: OpenWorkbook

        /// <summary>
        /// Speichert die Mappe unter ihrem Namen. Bei readOnly wird nicht gespeichert ( false ), sonst würde
        /// Excel scheitern oder eine unsichtbare Rückfrage zeigen.
        /// </summary>
        /// <returns>Erfolg</returns>
        public bool SaveWorkbook( )
        {
            if ( readOnly )
            {   // nur zum Lesen geöffnet
                Console.WriteLine( "SaveWorkbook: Mappe ist nur zum Lesen geöffnet, nicht gespeichert." );
                return ( false );
            }   // ende: nur zum Lesen geöffnet
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
        /// Der Name wird vervollständigt ( .xlsx, voller Pfad ); eine vorhandene Datei gleichen Namens wird
        /// ohne Rückfrage überschrieben ( Excels Rückfragen sind währenddessen aus, im unsichtbaren Excel
        /// würde sonst das Programm hängen ). Danach ist die Mappe beschreibbar ( readOnly = false ).
        /// </summary>
        /// <param name="newFileName"></param>
        /// <returns>boolean about the success of the operation</returns>
        public bool SaveWorkbookAs( string newFileName )
        {
            newFileName = CompleteTargetName( newFileName );
            bool alteRueckfragen = excelApp.DisplayAlerts;
            try
            {
                excelApp.DisplayAlerts = false;
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
            finally
            {   // Rückfragen wie vorher
                excelApp.DisplayAlerts = alteRueckfragen;
            }   // Ende: Rückfragen wie vorher

            fileName = newFileName;
            readOnly = false;
            return ( true );

        }   // end: SaveWorkbookAs

        /// <summary>
        /// Empty the current worksheet by clearing all its cells. 
        /// This method does not delete the worksheet itself, 
        /// but removes all data and formatting from it.
        /// </summary>
        public void CleanSheet( )
        {
            // das gewählte Blatt ( 27.11.26: vorher ActiveSheet.Worksheet, diese Eigenschaft gibt es nicht )
            worksheet = SelectedSheet( );
            worksheet.Cells.Clear();

        }   // end: CleanSheet

        /* ------------------------------ Blätter ( 27.11.26 )
         *  - Nummern sind 0-basiert ( erstes Blatt = 0 ), Excel zählt ab 1: Umrechnung hier, wie bei den Zellen
         *  - gezählt werden nur Arbeitsblätter ( Worksheets ), keine Diagrammblätter
         *  - sheetNumber ( aus ReadSheetNames, SelectSheet oder NewSheet ) bestimmt das Blatt für Lesen und
         *    Schreiben; sonst gilt das aktive Blatt dieser Mappe
        */

        /// <summary>
        /// Das Blatt für Lesen und Schreiben: das Blatt sheetNumber, sonst das aktive Blatt dieser Mappe.
        /// </summary>
        Excel.Worksheet SelectedSheet( )
        {
            if ( ( sheetNumber >= 0 ) && ( sheetNumber < workbook.Worksheets.Count ) )
                return ( (Excel.Worksheet)workbook.Worksheets[ sheetNumber + 1 ] );   // Excel zählt ab 1
            return ( (Excel.Worksheet)workbook.ActiveSheet );

        }   // Ende: Excel.Worksheet SelectedSheet

        /// <summary>
        /// Wählt das Blatt für die folgenden Lese- und Schreibaufrufe ohne Dialog.
        /// </summary>
        /// <param name="numTable">Nummer des Blatts, 0 = erstes</param>
        /// <returns>true, wenn es das Blatt gibt; sonst bleibt die Wahl unverändert</returns>
        public bool SelectSheet( int numTable )
        {
            if ( ( numTable < 0 ) || ( numTable >= workbook.Worksheets.Count ) )
                return ( false );
            sheetNumber = numTable;
            return ( true );

        }   // Ende: bool SelectSheet

        /// <summary>
        /// Returns the number of a chosen sheet. A dialog will open to let you choose from
        /// the found sheet names. Nummer 0-basiert; ohne Auswahl -1 ( dann gilt das aktive Blatt ).
        /// </summary>
        /// <returns>the number</returns>
        public int ReadSheetNames( )
        {
            // die Daten auslesen ( nur Arbeitsblätter, wie SelectedSheet )
            Excel.Sheets sheetsList = workbook.Worksheets;

            if ( sheetsList.Count > 0 )
            {
                List<string> locSheets = new List<string>();

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
        /// <param name="numTable">number of the sheet, 0 = erstes</param>
        /// <returns>the name or 'string.empty'</returns>
        public string GetSheetName( int numTable )
        {
            // Excel zählt ab 1 ( 27.11.26: vorher Sheets[ 0 ], "Ungültiger Index" bzw. falsches Blatt )
            Excel.Sheets sheetsList = workbook.Worksheets;

            if ( ( numTable >= 0 ) && ( numTable < sheetsList.Count ) )
            {
                return ( ( (Excel.Worksheet)sheetsList[ numTable + 1 ] ).Name );
            }
            return ( string.Empty );

        }   // end: GetSheetName

        /// <summary>
        /// Spaltennamen eines Blatts, ohne die Daten zu übernehmen ( 28.11.26, wie AccessOleDb und ExcelOleDb ).
        /// Bei useHeader die erste Zeile des Datenrechtecks, sonst die Spaltennamen des Blatts ( z. B. B, C, ... );
        /// dieselben Namen, die ReadStrings / ReadDoubles in 'headers' liefern. Die Blattwahl ( 'sheetNumber' )
        /// bleibt unverändert.
        /// </summary>
        /// <param name="numTable">Nummer des Blatts, 0 = erstes</param>
        /// <returns>die Spaltennamen</returns>
        /// <exception cref="InvalidOperationException">keine Mappe geöffnet</exception>
        /// <exception cref="ArgumentOutOfRangeException">kein Blatt mit dieser Nummer</exception>
        public string[] ReadColumnNames( int numTable )
        {
            if ( !workbookOpen )
                throw new InvalidOperationException( "ExcelInterop.ReadColumnNames: keine Mappe geöffnet ( " + lastError + " )." );
            if ( ( numTable < 0 ) || ( numTable >= workbook.Worksheets.Count ) )
                throw new ArgumentOutOfRangeException( nameof( numTable ),
                    $"ExcelInterop: Blatt Nr. {numTable} gibt es nicht ( {workbook.Worksheets.Count} Blätter in {fileName} )." );

            int altesBlatt = sheetNumber;
            try
            {   // Blatt vorübergehend wählen, Datenrechteck wie beim Lesen
                sheetNumber = numTable;
                object?[,] werte = ReadUsedRangeValues( out int firstCol, out int firstRow );
                CheckDataRectangle( werte, firstRow, firstCol );
                return ( ReadHeaders( werte, firstCol ) );
            }   // Ende: Blatt vorübergehend wählen, Datenrechteck wie beim Lesen
            finally
            {   // Blattwahl des Aufrufers wiederherstellen
                sheetNumber = altesBlatt;
            }   // Ende: Blattwahl des Aufrufers wiederherstellen

        }   // Ende: string[] ReadColumnNames

        /// <summary>
        /// Sets the name of a chosen sheet.
        /// </summary>
        /// <param name="numTable">number of the sheet, 0 = erstes</param>
        /// <param name="newName"></param>
        public void SetSheetName( int numTable, string newName )
        {
            // Excel zählt ab 1
            Excel.Sheets sheetsList = workbook.Worksheets;
            if ( ( numTable >= 0 ) && ( numTable < sheetsList.Count ) )
            {
                ( (Excel.Worksheet)sheetsList[ numTable + 1 ] ).Name = newName;
            }

        }   // end: SetSheetName

        /// <summary>
        /// Creates a new sheet at the end of the workbook and selects it ( sheetNumber ).
        /// Ohne Namen vergibt Excel einen ( ein leerer Name wäre ein Fehler ).
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns>the number of the new sheet, 0-basiert</returns>
        public int NewSheet( string sheetName = "" )
        {
            // ans Ende anhängen ( 27.11.26: vorher vor dem aktiven Blatt, Rückgabe ab 1, leerer Name = Fehler )
            Excel.Sheets sheetsList = workbook.Worksheets;
            Excel.Worksheet newSheet =
                (Excel.Worksheet)sheetsList.Add( After: sheetsList[ sheetsList.Count ] );
            if ( !string.IsNullOrWhiteSpace( sheetName ) )
                newSheet.Name = sheetName;

            sheetNumber = sheetsList.Count - 1;
            return ( sheetNumber );

        }   // end: NewSheet

        /* ------------------------------ Lesen ( 27.11.26 )
         *  - gelesen wird das Datenrechteck des aktiven Blatts: der benutzte Bereich ( UsedRange ), auch wenn er
         *    nicht in A1 beginnt ( vorher ab A1 gezählt, eine versetzte Tabelle verlor dadurch Daten )
         *  - alle Werte kommen mit einem Zugriff aus Excel ( Value2 als object[,] ), danach wird jeder Wert
         *    einzeln geprüft und umgewandelt ( vorher eine Abfrage je Zelle: 2000 x 5 Zellen 23 s, jetzt 12 ms )
         *  - leere Zellen ergeben Standardwerte ( "" bzw. 0 )
        */

        /// <summary>
        /// Reads the table's data types as 2d-array of
        /// strings into 'valuesTypesArray'. Needed for the analysis
        /// of foreign data. Ganzes Datenrechteck einschließlich Kopfzeile; leere Zelle: leerer Text.
        /// </summary>
        public void ReadTypes( )
        {
            // zuerst leeren, lokal füllen, am Ende zuweisen ( wie ReadDoubles, 27.11.26 ): nach einem Fehler
            // keine Typen eines früher gelesenen Blatts
            valuesTypesArray = null!;

            // alle Werte des Datenrechtecks mit einem Zugriff ( Abschnitt "Lesen" );
            // ohne Prüfung des Rechtecks: ReadTypes dient der Untersuchung fremder Daten
            object?[,] werte = ReadUsedRangeValues( out int firstCol, out int firstRow );
            int rowCount = werte.GetLength( 0 );
            int colCount = werte.GetLength( 1 );

            string[,] locTypesArray = new string[ colCount, rowCount ];

            for ( int rowInd = 1; rowInd <= rowCount; rowInd++ )
            {
                for ( int colInd = 1; colInd <= colCount; colInd++ )
                {
                    // Typ des Zellwerts; Excel zählt ab 1, das Array ab 0
                    object? wert = werte[ rowInd, colInd ];
                    locTypesArray[ colInd - 1, rowInd - 1 ] = wert?.GetType().ToString() ?? string.Empty;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

            valuesTypesArray = locTypesArray;

        }   // end: ReadTypes

        /// <summary>
        /// Reads the table's data as 2d-array of
        /// strings into 'valuesStringArray' ( [ spalte, zeile ] ).
        /// Bei useHeader steht die erste Zeile in 'headers' und nicht in den Daten ( wie ExcelOleDb ).
        /// Excel-Fehler erscheinen als Fehlertext ( #DIV/0!, #NV, ... ).
        /// </summary>
        public void ReadStrings()
        {
            // zuerst leeren, lokal füllen, am Ende zuweisen ( wie ReadDoubles, 27.11.26 ): Scheitert das Lesen
            // ( z. B. leeres Blatt ), bleiben keine Texte eines früher gelesenen Blatts stehen; die
            // Schreibfunktionen melden dann "keine Texte vorhanden" ( CheckStrings ); die Kopfzeile ebenso
            valuesStringArray = null!;
            headersString = new string[ 0 ];
            headers = new string[ 0 ];

            // alle Werte des Datenrechtecks mit einem Zugriff ( Abschnitt "Lesen" )
            object?[,] werte = ReadUsedRangeValues( out int firstCol, out int firstRow );
            // die Daten haben zu passen, sonst Fehlermeldung ( CheckDataRectangle )
            CheckDataRectangle( werte, firstRow, firstCol );
            int rowCount = werte.GetLength( 0 );
            int colCount = werte.GetLength( 1 );

            /* Kopfzeile wie in ExcelOleDb ( HDR=YES ):
             *  - useHeader: erste Zeile des Datenrechtecks nach 'headers', die Daten beginnen in der zweiten
             *  - sonst: alle Zeilen sind Daten, 'headers' enthält die Spaltennamen des Blatts ( z. B. B, C, ... )
            */
            string[] locHeaders = ReadHeaders( werte, firstCol );
            int startRow = useHeader ? 2 : 1;
            CheckDataRows( rowCount, startRow, "ReadStrings" );

            string[,] locStringArray = new string[ colCount, Math.Max( 0, rowCount - startRow + 1 ) ];

            for ( int rowInd = startRow; rowInd <= rowCount; rowInd++ )
            {
                for ( int colInd = 1; colInd <= colCount; colInd++ )
                {
                    // jeder Wert einzeln geprüft; Excel zählt ab 1, das Array ab 0 ( Datenzeilen ab startRow )
                    locStringArray[ colInd - 1, rowInd - startRow ] = CellText( werte[ rowInd, colInd ] );

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

            valuesStringArray = locStringArray;
            headersString = locHeaders;
            headers = locHeaders;

        }   // end: ReadStrings

        /// <summary>
        /// Liefert die Kopfzeile für 'headers': bei useHeader die Texte der ersten Zeile des Datenrechtecks,
        /// sonst die Spaltennamen des Blatts ( das Rechteck kann z. B. in Spalte B beginnen ).
        /// Setzt 'headers' nicht selbst: der Aufrufer weist es erst nach fehlerfreiem Lesen zu ( 27.11.26 ).
        /// </summary>
        /// <param name="werte">Werte des Datenrechtecks [ zeile, spalte ], ab 1</param>
        /// <param name="firstCol">erste Spalte des Rechtecks im Blatt ( ab 1 )</param>
        /// <returns>je Spalte ein Name</returns>
        string[] ReadHeaders( object?[,] werte, int firstCol )
        {
            int colCount = werte.GetLength( 1 );
            string[] locHeaders = new string[ colCount ];
            for ( int colInd = 1; colInd <= colCount; colInd++ )
            {   // je Spalte
                if ( useHeader )
                    locHeaders[ colInd - 1 ] = CellText( werte[ 1, colInd ] );
                else
                    locHeaders[ colInd - 1 ] = GetColumnFromNumber( firstCol + colInd - 1 );
            }   // Ende: je Spalte
            return ( locHeaders );

        }   // Ende: string[] ReadHeaders

        /// <summary>
        /// Holt die Werte des Datenrechtecks ( benutzter Bereich des aktiven Blatts ) mit einem Zugriff.
        /// </summary>
        /// <param name="firstCol">erste Spalte des Rechtecks im Blatt ( ab 1 )</param>
        /// <param name="firstRow">erste Zeile des Rechtecks im Blatt ( ab 1 )</param>
        /// <returns>Werte [ zeile, spalte ], ab 1 gezählt wie in Excel</returns>
        object?[,] ReadUsedRangeValues( out int firstCol, out int firstRow )
        {
            // das gewählte Blatt ( 27.11.26: vorher ActiveSheet.Worksheet, diese Eigenschaft gibt es nicht )
            worksheet = SelectedSheet( );
            Excel.Range usedRange = worksheet.UsedRange;
            firstCol = usedRange.Column;
            firstRow = usedRange.Row;
            return ( RangeValues( usedRange ) );

        }   // Ende: object?[,] ReadUsedRangeValues

        /// <summary>
        /// Werte eines Bereichs mit einem Zugriff ( Value2 ). Eine einzelne Zelle liefert in Excel keinen
        /// Array, sie wird in ein Feld 1 x 1 gelegt ( ab 1 wie Excel ).
        /// </summary>
        /// <param name="range">der Bereich</param>
        /// <returns>Werte [ zeile, spalte ], ab 1 gezählt</returns>
        static object?[,] RangeValues( Excel.Range range )
        {
            object? werte = range.Value2;
            if ( werte is object[,] feld )
                return ( feld );
            object?[,] einzeln = (object?[,])Array.CreateInstance(
                typeof( object ), new[] { 1, 1 }, new[] { 1, 1 } );
            einzeln[ 1, 1 ] = werte;
            return ( einzeln );

        }   // Ende: object?[,] RangeValues

        /// <summary>
        /// Zellwert als Text: leer = "", Excel-Fehler als Fehlertext ( #DIV/0!, #NV, ... ),
        /// Zahlen in der Schreibweise von Windows ( z. B. 1,5 ).
        /// </summary>
        static string CellText( object? wert )
        {
            if ( wert == null )
                return ( string.Empty );
            // Value2 liefert Zahlen als double, einen int nur für Excel-Fehler
            if ( wert is int code )
                return ( ErrorText( code ) );
            return ( wert.ToString() ?? string.Empty );

        }   // Ende: string CellText

        /// <summary>
        /// Zellwert als Zahl: Zahl ( auch Datum als Seriennummer ) unverändert, WAHR = 1 / FALSCH = 0,
        /// Text mit einer Zahl in der Schreibweise von Windows ( z. B. "1,5" ) umgewandelt; leer ( auch nur
        /// Leerzeichen ) und Excel-Fehler 0. Anderer Text ist keine Zahl: false ( 27.11.26, vorher still 0 ).
        /// </summary>
        /// <param name="wert">Zellwert aus Value2</param>
        /// <param name="zahl">die Zahl</param>
        /// <returns>false bei Text, der keine Zahl ist</returns>
        static bool TryCellNumber( object? wert, out double zahl )
        {
            zahl = 0.0;
            switch ( wert )
            {   // je nach Typ des Zellwerts
                case double wertZahl:
                    zahl = wertZahl;
                    return ( true );
                case bool wahr:
                    zahl = wahr ? 1.0 : 0.0;
                    return ( true );
                case string text:
                    if ( string.IsNullOrWhiteSpace( text ) )
                        return ( true );   // leer: Standardwert 0
                    return ( double.TryParse( text.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture,
                        out zahl ) );
                default:
                    return ( true );       // leer ( null ) oder Excel-Fehler: 0
            }   // Ende: je nach Typ des Zellwerts

        }   // Ende: bool TryCellNumber

        /// <summary>
        /// Fehlermeldung für eine Zelle mit Text statt Zahl ( die Daten haben zu passen ).
        /// </summary>
        /// <param name="wert">Zellwert</param>
        /// <param name="spalte">Spalte im Blatt, ab 1</param>
        /// <param name="zeile">Zeile im Blatt, ab 1</param>
        /// <param name="funktion">aufrufende Funktion</param>
        /// <returns>die Ausnahme zum Werfen</returns>
        InvalidDataException TextInNumberCell( object? wert, int spalte, int zeile, string funktion )
        {
            return ( new InvalidDataException( $"ExcelInterop.{funktion}: Zelle {GetColumnFromNumber( spalte )}{zeile} "
                + $"enthält den Text '{wert}' statt einer Zahl ( nur leere Zellen werden 0 )." ) );

        }   // Ende: InvalidDataException TextInNumberCell

        /// <summary>
        /// Die Daten haben zu passen: Unter der Kopfzeile muss es Daten geben ( 27.11.26; ein Blatt nur mit
        /// Kopfzeile ist wie ein leeres Blatt kein Datensatz, vorher leeres Ergebnis ohne Meldung ).
        /// </summary>
        /// <param name="rowCount">Zeilen des Datenrechtecks</param>
        /// <param name="startRow">erste Datenzeile im Rechteck ( 2 mit Kopfzeile )</param>
        /// <param name="funktion">aufrufende Funktion</param>
        /// <exception cref="InvalidDataException">keine Datenzeile</exception>
        void CheckDataRows( int rowCount, int startRow, string funktion )
        {
            if ( rowCount < startRow )
                throw new InvalidDataException( $"ExcelInterop.{funktion}: Blatt '{worksheet.Name}' enthält unter der "
                    + "Kopfzeile keine Daten." );

        }   // Ende: void CheckDataRows

        /// <summary>
        /// Fehlertext zu einem Excel-Fehlercode aus Value2 ( Schreibweise des deutschen Excel ).
        /// </summary>
        static string ErrorText( int code )
        {
            switch ( code )
            {   // Excel-Fehlercodes ( CVErr )
                case -2146826288: return ( "#NULL!" );
                case -2146826281: return ( "#DIV/0!" );
                case -2146826273: return ( "#WERT!" );
                case -2146826265: return ( "#BEZUG!" );
                case -2146826259: return ( "#NAME?" );
                case -2146826252: return ( "#ZAHL!" );
                case -2146826246: return ( "#NV" );
                default: return ( "#FEHLER " + code );
            }   // Ende: Excel-Fehlercodes ( CVErr )

        }   // Ende: string ErrorText

        /// <summary>
        /// Leere Zelle: kein Wert oder leerer Text.
        /// </summary>
        static bool IsEmptyCell( object? wert )
        {
            return ( ( wert == null ) || ( ( wert is string text ) && ( text.Length == 0 ) ) );

        }   // Ende: bool IsEmptyCell

        /* Die Daten haben zu passen, sonst Fehlermeldung ( Vorgabe vom 27.11.26 )
         *  - ein Datenrechteck ohne jeden Inhalt ( leeres Blatt ) ist kein Datensatz
         *  - eine ganz leere erste oder letzte Zeile bzw. Spalte gehört nicht zu den Daten, sondern meist zu
         *    Formatierung oder Resten außerhalb ( Excel zählt sie zum benutzten Bereich ); sie würde sonst als
         *    Muster aus lauter Standardwerten gelesen
         *  - einzelne leere Zellen innerhalb der Daten bekommen weiter Standardwerte
        */

        /// <summary>
        /// Prüft das Datenrechteck: nicht leer, keine ganz leere Rand-Zeile oder -Spalte.
        /// </summary>
        /// <param name="werte">Werte [ zeile, spalte ], ab 1</param>
        /// <param name="firstRow">erste Zeile des Rechtecks im Blatt</param>
        /// <param name="firstCol">erste Spalte des Rechtecks im Blatt</param>
        /// <exception cref="InvalidDataException">das Rechteck passt nicht</exception>
        void CheckDataRectangle( object?[,] werte, int firstRow, int firstCol )
        {
            int rowCount = werte.GetLength( 0 );
            int colCount = werte.GetLength( 1 );

            bool ZeileLeer( int zeile )
            {   // ganze Zeile ohne Inhalt?
                for ( int spalte = 1; spalte <= colCount; spalte++ )
                    if ( !IsEmptyCell( werte[ zeile, spalte ] ) )
                        return ( false );
                return ( true );
            }   // Ende: ZeileLeer
            bool SpalteLeer( int spalte )
            {   // ganze Spalte ohne Inhalt?
                for ( int zeile = 1; zeile <= rowCount; zeile++ )
                    if ( !IsEmptyCell( werte[ zeile, spalte ] ) )
                        return ( false );
                return ( true );
            }   // Ende: SpalteLeer

            string blatt = "Blatt '" + worksheet.Name + "'";
            if ( ( rowCount == 1 ) && ( colCount == 1 ) && IsEmptyCell( werte[ 1, 1 ] ) )
                throw new InvalidDataException( "ExcelInterop: " + blatt + " enthält keine Daten." );
            if ( ZeileLeer( 1 ) )
                throw new InvalidDataException( "ExcelInterop: " + blatt + ", Zeile " + firstRow
                    + " ist leer, gehört aber zum Datenrechteck ( Formatierung oder Reste entfernen )." );
            if ( ZeileLeer( rowCount ) )
                throw new InvalidDataException( "ExcelInterop: " + blatt + ", Zeile " + ( firstRow + rowCount - 1 )
                    + " ist leer, gehört aber zum Datenrechteck ( Formatierung oder Reste entfernen )." );
            if ( SpalteLeer( 1 ) )
                throw new InvalidDataException( "ExcelInterop: " + blatt + ", Spalte " + GetColumnFromNumber( firstCol )
                    + " ist leer, gehört aber zum Datenrechteck ( Formatierung oder Reste entfernen )." );
            if ( SpalteLeer( colCount ) )
                throw new InvalidDataException( "ExcelInterop: " + blatt + ", Spalte "
                    + GetColumnFromNumber( firstCol + colCount - 1 )
                    + " ist leer, gehört aber zum Datenrechteck ( Formatierung oder Reste entfernen )." );

        }   // Ende: void CheckDataRectangle

        /// <summary>
        /// Prüft eine Matrix( x, y ) vor dem Schreiben: NaN und Unendlich kann Excel nicht speichern
        /// ( es schreibt #ZAHL! ), daher Fehlermeldung statt unbrauchbarer Zellen.
        /// </summary>
        /// <exception cref="ArgumentException">NaN oder Unendlich in der Matrix</exception>
        /// <exception cref="InvalidOperationException">keine Matrix ( z. B. ReadDoubles gescheitert )</exception>
        static void CheckFinite( double[,] matrix, string funktion )
        {
            // keine Daten: ReadDoubles leert 'valuesDoubleArray' zu Beginn ( 27.11.26 ), nach einem Fehler bleibt es leer
            if ( matrix == null )
                throw new InvalidOperationException( "ExcelInterop." + funktion
                    + ": keine Zahlen vorhanden ( zuerst ReadDoubles fehlerfrei ausführen )." );
            for ( int x = 0; x < matrix.GetLength( 0 ); x++ )
                for ( int y = 0; y < matrix.GetLength( 1 ); y++ )
                    if ( !double.IsFinite( matrix[ x, y ] ) )
                        throw new ArgumentException( "ExcelInterop." + funktion + ": Wert an ( " + x + ", " + y
                            + " ) ist " + matrix[ x, y ] + ", Excel kann nur endliche Zahlen speichern." );

        }   // Ende: void CheckFinite

        /// <summary>
        /// Prüft vor dem Schreiben, ob es Texte gibt ( Gegenstück zu CheckFinite für Zahlen ).
        /// </summary>
        /// <exception cref="InvalidOperationException">keine Texte ( z. B. ReadStrings gescheitert )</exception>
        static void CheckStrings( string[,] matrix, string funktion )
        {
            // keine Daten: ReadStrings leert 'valuesStringArray' zu Beginn ( 27.11.26 ), nach einem Fehler bleibt es leer
            if ( matrix == null )
                throw new InvalidOperationException( "ExcelInterop." + funktion
                    + ": keine Texte vorhanden ( zuerst ReadStrings fehlerfrei ausführen )." );

        }   // Ende: void CheckStrings

        /// <summary>
        /// Prüft ein Feld vor dem Schreiben auf NaN und Unendlich ( wie CheckFinite für Matrizen ).
        /// </summary>
        /// <exception cref="ArgumentException">NaN oder Unendlich im Feld</exception>
        /// <exception cref="InvalidOperationException">kein Feld übergeben</exception>
        static void CheckFinite( double[] feld, string funktion )
        {
            if ( feld == null )
                throw new InvalidOperationException( "ExcelInterop." + funktion + ": keine Zahlen übergeben." );
            for ( int x = 0; x < feld.Length; x++ )
                if ( !double.IsFinite( feld[ x ] ) )
                    throw new ArgumentException( "ExcelInterop." + funktion + ": Wert an ( " + x + " ) ist "
                        + feld[ x ] + ", Excel kann nur endliche Zahlen speichern." );

        }   // Ende: void CheckFinite

        /// <summary>
        /// Vervollständigt einen Zielnamen: fehlende Endung wird ".xlsx", ein relativer Name wird zum vollen
        /// Pfad ( Programmverzeichnis ). So prüft File.Exists dieselbe Datei, die Excel speichert.
        /// </summary>
        static string CompleteTargetName( string name )
        {
            if ( string.IsNullOrWhiteSpace( name ) )
                return ( name );
            if ( Path.GetExtension( name ) == "" )
                name += ".xlsx";
            return ( Path.GetFullPath( name ) );

        }   // Ende: string CompleteTargetName

        /// <summary>
        /// Writes the 2d-array to the active sheet.
        /// Overwrites anything in its way and starts at A1.
        /// </summary>
        public void WriteStrings()
        {
            // die Daten haben zu passen: es muss Texte geben ( ReadStrings leert das Feld zu Beginn )
            CheckStrings( valuesStringArray, "WriteStrings" );

            // das gewählte Blatt ( 27.11.26: vorher nicht gesetzt, ohne vorheriges Lesen NullReferenceException )
            worksheet = SelectedSheet( );

            // calculate the Excel range to write to
            int colCount = valuesStringArray.GetLength( 0 );
            int rowCount = valuesStringArray.GetLength( 1 );
            if ( ( colCount == 0 ) || ( rowCount == 0 ) )
                return;   // nichts zu schreiben

            int startCol = 1;
            int endCol = startCol + colCount - 1;
            int startRow = 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startCol );
            string xDim2 = GetColumnFromNumber( endCol );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            // send the data ( Excel erwartet [ zeile, spalte ], siehe ToExcelOrder )
            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = ToExcelOrder( valuesStringArray );

        } // end: WriteStrings

        /// <summary>
        /// Übergabe an Excel: Die Klasse hält jede Matrix als Matrix( x, y ) wie Punkt( x, y ), also
        /// [ spalte, zeile ]. Excel erwartet bei Value2 dagegen [ zeile, spalte ] ( geprüft 27.11.26 ).
        /// Deshalb wird nur hier, unmittelbar vor dem Schreiben, in eine Kopie in Excels Reihenfolge umgesetzt.
        /// </summary>
        /// <param name="matrix">Matrix( x, y )</param>
        /// <returns>Kopie [ y, x ] für Range.Value2</returns>
        static T[,] ToExcelOrder<T>( T[,] matrix )
        {
            int colCount = matrix.GetLength( 0 );
            int rowCount = matrix.GetLength( 1 );
            T[,] excelFeld = new T[ rowCount, colCount ];
            for ( int x = 0; x < colCount; x++ )
                for ( int y = 0; y < rowCount; y++ )
                    excelFeld[ y, x ] = matrix[ x, y ];
            return ( excelFeld );

        }   // Ende: T[,] ToExcelOrder

        /// <summary>
        /// Reads a subset of the table's data as a 2D array of strings.
        /// Koordinaten im Blatt ab 1, X = Spalte, Y = Zeile, Anfang und Ende eingeschlossen.
        /// </summary>
        /// <param name="startX">upper left X coordinate</param>
        /// <param name="startY">upper left Y coordinate</param>
        /// <param name="endX">lower right X coordinate</param>
        /// <param name="endY">lower right Y coordinate</param>
        /// <returns>subset of the table's data as a 2D array of strings</returns>
        /// <exception cref="ArgumentException">Koordinate unter 1 oder Ende vor dem Anfang</exception>
        public string[,] ReadStrings( int startX, int startY, int endX, int endY )
        {
            // Bereich prüfen: Excel zählt ab 1, Ende nicht vor dem Anfang ( sonst still leeres Ergebnis )
            if ( ( startX < 1 ) || ( startY < 1 ) || ( endX < startX ) || ( endY < startY ) )
                throw new ArgumentException(
                    "ExcelInterop.ReadStrings: Bereich ungültig ( ab 1, Ende nicht vor dem Anfang )" );

            // das gewählte Blatt ( 27.11.26: vorher ActiveSheet.Worksheet, diese Eigenschaft gibt es nicht )
            worksheet = SelectedSheet( );

            // alle Werte des Bereichs mit einem Zugriff ( Abschnitt "Lesen" ); Cells[ Zeile, Spalte ]
            Excel.Range bereich = worksheet.Range[
                worksheet.Cells[ startY, startX ], worksheet.Cells[ endY, endX ] ];
            object?[,] werte = RangeValues( bereich );

            string[,] locStringArray = new string[ endX - startX + 1, endY - startY + 1 ];

            // a dream would be the next line, but it is not save for input
            // valuesStringArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as string[,];

            for ( int rowInd = 1; rowInd <= werte.GetLength( 0 ); rowInd++ )
            {
                for ( int colInd = 1; colInd <= werte.GetLength( 1 ); colInd++ )
                {
                    // jeder Wert einzeln geprüft; Excel zählt ab 1, das Array ab 0
                    locStringArray[ colInd - 1, rowInd - 1 ] = CellText( werte[ rowInd, colInd ] );

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

            return( locStringArray );

        }   // end: ReadStrings

        /// <summary>
        /// Writes the 2d-array to the active sheet at the specified coordinates.
        /// Overwrites anything in its way. Koordinaten im Blatt ab 1, X = Spalte, Y = Zeile.
        /// </summary>
        /// <exception cref="ArgumentException">Koordinate unter 1</exception>
        public void WriteStrings( string[,] dataArray, int startX, int startY )
        {
            // Koordinaten prüfen: Excel zählt ab 1
            if ( ( startX < 1 ) || ( startY < 1 ) )
                throw new ArgumentException( "ExcelInterop.WriteStrings: Koordinaten ab 1" );

            // das gewählte Blatt ( 27.11.26: vorher nicht gesetzt, ohne vorheriges Lesen NullReferenceException )
            worksheet = SelectedSheet( );

            // calculate the Excel range to write to
            int colCount = dataArray.GetLength( 0 );
            int rowCount = dataArray.GetLength( 1 );
            if ( ( colCount == 0 ) || ( rowCount == 0 ) )
                return;   // nichts zu schreiben

            int startCol = startX;
            int endCol = startCol + colCount - 1;
            int startRow = startY;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startCol );
            string xDim2 = GetColumnFromNumber( endCol );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            // Excel erwartet [ zeile, spalte ], siehe ToExcelOrder
            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = ToExcelOrder( dataArray );

        } // end: WriteStrings

        /// <summary>
        /// Appends the 2d-array to the active sheet after the last row.
        /// Unter dem Datenrechteck, ab dessen erster Spalte; auf einem leeren Blatt ab A1.
        /// </summary>
        /// <param name="dataArray">2d string array, Matrix( x, y )</param>
        public void AppendRowStringMatrix(string[,] dataArray)
        {
            // erste freie Zeile unter dem Datenrechteck ( 27.11.26: vorher SpecialCells( xlCellTypeLastCell ),
            // das merkt sich die frühere letzte Zeile auch nach dem Löschen )
            NextFreeRow( out int startRow, out int startCol );

            // calculate the Excel range to write to
            int colCount = dataArray.GetLength(0);
            int rowCount = dataArray.GetLength(1);
            if ( ( colCount == 0 ) || ( rowCount == 0 ) )
                return;   // nichts zu schreiben

            int endCol = startCol + colCount - 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber(startCol);
            string xDim2 = GetColumnFromNumber(endCol);
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            // Excel erwartet [ zeile, spalte ], siehe ToExcelOrder
            worksheet.Range[xDim1 + yDim1, xDim2 + yDim2].Value2 = ToExcelOrder( dataArray );

        } // end: AppendRowStringMatrix

        /// <summary>
        /// Writes the array to the active sheet after the last row.
        /// Unter dem Datenrechteck, ab dessen erster Spalte; auf einem leeren Blatt ab A1.
        /// </summary>
        /// <param name="dataArray">string array ( eine Zeile )</param>
        public void AppendRowStringField(string[] dataArray)
        {
            // erste freie Zeile unter dem Datenrechteck ( 27.11.26: vorher SpecialCells( xlCellTypeLastCell ),
            // das merkt sich die frühere letzte Zeile auch nach dem Löschen )
            NextFreeRow( out int startRow, out int startCol );

            // calculate the Excel range to write to
            int colCount = dataArray.Length;
            if ( colCount == 0 )
                return;   // nichts zu schreiben

            int endCol = startCol + colCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber(startCol);
            string xDim2 = GetColumnFromNumber(endCol);
            string yDim1 = startRow.ToString();

            // ein eindimensionales Feld schreibt Excel als eine Zeile ( keine Umsetzung nötig )
            worksheet.Range[xDim1 + yDim1, xDim2 + yDim1].Value2 = dataArray;

        } // end: AppendRowStringField

        /// <summary>
        /// Erste freie Zeile unter dem Datenrechteck des aktiven Blatts und dessen erste Spalte;
        /// auf einem leeren Blatt Zeile 1, Spalte 1. Setzt 'worksheet'.
        /// </summary>
        /// <param name="startRow">erste freie Zeile ( ab 1 )</param>
        /// <param name="startCol">erste Spalte des Datenrechtecks ( ab 1 )</param>
        void NextFreeRow( out int startRow, out int startCol )
        {
            // das gewählte Blatt ( 27.11.26: vorher nicht gesetzt, ohne vorheriges Lesen NullReferenceException )
            worksheet = SelectedSheet( );
            if ( excelApp.WorksheetFunction.CountA( worksheet.Cells ) == 0 )
            {   // leeres Blatt
                startRow = 1;
                startCol = 1;
                return;
            }   // ende: leeres Blatt
            Excel.Range usedRange = worksheet.UsedRange;
            // die Daten haben zu passen: keine leere Rand-Zeile oder -Spalte ( sonst Lücke beim Anhängen )
            CheckDataRectangle( RangeValues( usedRange ), usedRange.Row, usedRange.Column );
            startRow = usedRange.Row + usedRange.Rows.Count;
            startCol = usedRange.Column;

        }   // Ende: void NextFreeRow

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// doubles into 'valuesDoubleArray' ( [ spalte, zeile ] ). Umwandlung je Zelle siehe TryCellNumber
        /// ( leer, Excel-Fehler: 0; Text ohne Zahl: InvalidDataException mit Zelladresse ).
        /// Bei useHeader steht die erste Zeile in 'headers' und nicht in den Daten ( wie ReadStrings ).
        /// </summary>
        public void ReadDoubles()
        {
            // zuerst leeren ( Vorgabe vom 27.11.26 ): Scheitert dieses Lesen, bleiben keine Zahlen eines früher
            // gelesenen Blatts stehen, die sich ohne Hinweis speichern ließen; die Schreibfunktionen melden dann
            // "keine Zahlen vorhanden" ( CheckFinite ); die Kopfzeile ebenso
            valuesDoubleArray = null!;
            headersDouble = new string[ 0 ];
            headers = new string[ 0 ];

            // alle Werte des Datenrechtecks mit einem Zugriff ( Abschnitt "Lesen" )
            object?[,] werte = ReadUsedRangeValues( out int firstCol, out int firstRow );
            // die Daten haben zu passen, sonst Fehlermeldung ( CheckDataRectangle )
            CheckDataRectangle( werte, firstRow, firstCol );
            int rowCount = werte.GetLength( 0 );
            int colCount = werte.GetLength( 1 );

            // Kopfzeile wie in ReadStrings ( useHeader: erste Zeile nach 'headers', Daten ab der zweiten )
            string[] locHeaders = ReadHeaders( werte, firstCol );
            int startRow = useHeader ? 2 : 1;
            CheckDataRows( rowCount, startRow, "ReadDoubles" );

            // erst lokal füllen, erst am Ende zuweisen ( wie ReadDoubles für einen Bereich ): Bricht das Lesen
            // bei Text ab, bleibt kein halb gefülltes Feld zurück ( 27.11.26: vorher Rest 0, so auch speicherbar )
            double[,] locDoubleArray = new double[ colCount, Math.Max( 0, rowCount - startRow + 1 ) ];

            // a dream would be the next line, but it is not save for input
            // valuesDoubleArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as double[,];

            for (int rowInd = startRow; rowInd <= rowCount; rowInd++)
            {
                for (int colInd = 1; colInd <= colCount; colInd++)
                {
                    // jeder Wert einzeln geprüft; Excel zählt ab 1, das Array ab 0 ( Datenzeilen ab startRow )
                    if ( !TryCellNumber( werte[ rowInd, colInd ], out double zahl ) )
                        throw TextInNumberCell( werte[ rowInd, colInd ], firstCol + colInd - 1,
                            firstRow + rowInd - 1, "ReadDoubles" );
                    locDoubleArray[ colInd - 1, rowInd - startRow ] = zahl;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

            valuesDoubleArray = locDoubleArray;
            headersDouble = locHeaders;
            headers = locHeaders;

        }   // end: ReadDoubles

        /// <summary>
        /// Writes the 2d-array to the active sheet.
        /// Overwrites anything in its way and starts at A1.
        /// </summary>
        public void WriteDoubles( )
        {
            // die Daten haben zu passen: nur endliche Zahlen ( sonst Fehlermeldung statt #ZAHL! )
            CheckFinite( valuesDoubleArray, "WriteDoubles" );

            // das gewählte Blatt ( 27.11.26: vorher nicht gesetzt, ohne vorheriges Lesen NullReferenceException )
            worksheet = SelectedSheet( );

            // calculate the Excel range to write to
            int colCount = valuesDoubleArray.GetLength( 0 );
            int rowCount = valuesDoubleArray.GetLength( 1 );
            if ( ( colCount == 0 ) || ( rowCount == 0 ) )
                return;   // nichts zu schreiben

            int startCol = 1;
            int endCol = startCol + colCount - 1;
            int startRow = 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startCol );
            string xDim2 = GetColumnFromNumber( endCol );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            // Excel erwartet [ zeile, spalte ], siehe ToExcelOrder
            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = ToExcelOrder( valuesDoubleArray );

        } // end: WriteDoubles

        /// <summary>
        /// Reads a subset of the table's data as a 2D array of doubles.
        /// Koordinaten im Blatt ab 1, X = Spalte, Y = Zeile, Anfang und Ende eingeschlossen.
        /// </summary>
        /// <param name="startX">upper left X coordinate</param>
        /// <param name="startY">upper left Y coordinate</param>
        /// <param name="endX">lower right X coordinate</param>
        /// <param name="endY">lower right Y coordinate</param>
        /// <returns>subset of the table's data as a 2D array of doubles</returns>
        /// <exception cref="ArgumentException">Koordinate unter 1 oder Ende vor dem Anfang</exception>
        public double[,] ReadDoubles( int startX, int startY, int endX, int endY )
        {
            // Bereich prüfen: Excel zählt ab 1, Ende nicht vor dem Anfang ( sonst still leeres Ergebnis )
            if ( ( startX < 1 ) || ( startY < 1 ) || ( endX < startX ) || ( endY < startY ) )
                throw new ArgumentException(
                    "ExcelInterop.ReadDoubles: Bereich ungültig ( ab 1, Ende nicht vor dem Anfang )" );

            // das gewählte Blatt ( 27.11.26: vorher ActiveSheet.Worksheet, diese Eigenschaft gibt es nicht )
            worksheet = SelectedSheet( );

            // alle Werte des Bereichs mit einem Zugriff ( Abschnitt "Lesen" ); Cells[ Zeile, Spalte ]
            Excel.Range bereich = worksheet.Range[
                worksheet.Cells[ startY, startX ], worksheet.Cells[ endY, endX ] ];
            object?[,] werte = RangeValues( bereich );

            double[,] locDoubleArray = new double[ endX - startX + 1, endY - startY + 1 ];

            // a dream would be the next line, but it is not save for input
            // valuesDoubleArray = worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 as double[,];

            for ( int rowInd = 1; rowInd <= werte.GetLength( 0 ); rowInd++ )
            {
                for ( int colInd = 1; colInd <= werte.GetLength( 1 ); colInd++ )
                {
                    // jeder Wert einzeln geprüft; Excel zählt ab 1, das Array ab 0
                    if ( !TryCellNumber( werte[ rowInd, colInd ], out double zahl ) )
                        throw TextInNumberCell( werte[ rowInd, colInd ], startX + colInd - 1,
                            startY + rowInd - 1, "ReadDoubles" );
                    locDoubleArray[ colInd - 1, rowInd - 1 ] = zahl;

                }   // end: for ( int colInd

            }   // end: for ( int rowInd

            return ( locDoubleArray );

        }   // end: ReadDoubles

        /// <summary>
        /// Writes the 2d-array to the active sheet at the specified coordinates.
        /// Overwrites anything in its way. Koordinaten im Blatt ab 1, X = Spalte, Y = Zeile.
        /// </summary>
        /// <exception cref="ArgumentException">Koordinate unter 1</exception>
        public void WriteDoubles( double[,] dataArray, int startX, int startY )
        {
            // Koordinaten prüfen: Excel zählt ab 1
            if ( ( startX < 1 ) || ( startY < 1 ) )
                throw new ArgumentException( "ExcelInterop.WriteDoubles: Koordinaten ab 1" );
            // die Daten haben zu passen: nur endliche Zahlen ( sonst Fehlermeldung statt #ZAHL! )
            CheckFinite( dataArray, "WriteDoubles" );

            // das gewählte Blatt ( 27.11.26: vorher nicht gesetzt, ohne vorheriges Lesen NullReferenceException )
            worksheet = SelectedSheet( );

            // calculate the Excel range to write to
            int colCount = dataArray.GetLength( 0 );
            int rowCount = dataArray.GetLength( 1 );
            if ( ( colCount == 0 ) || ( rowCount == 0 ) )
                return;   // nichts zu schreiben

            int startCol = startX;
            int endCol = startCol + colCount - 1;
            int startRow = startY;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber( startCol );
            string xDim2 = GetColumnFromNumber( endCol );
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            // Excel erwartet [ zeile, spalte ], siehe ToExcelOrder
            worksheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = ToExcelOrder( dataArray );

        } // end: WriteDoubles

        /// <summary>
        /// Appends the 2d-array to the active sheet after the last row.
        /// Unter dem Datenrechteck, ab dessen erster Spalte; auf einem leeren Blatt ab A1.
        /// </summary>
        /// <param name="dataArray">2d double array, Matrix( x, y )</param>
        public void AppendRowDoublesMatrix(double[,] dataArray)
        {
            // die Daten haben zu passen: nur endliche Zahlen ( sonst Fehlermeldung statt #ZAHL! )
            CheckFinite( dataArray, "AppendRowDoublesMatrix" );

            // erste freie Zeile unter dem Datenrechteck ( 27.11.26: vorher SpecialCells( xlCellTypeLastCell ),
            // das merkt sich die frühere letzte Zeile auch nach dem Löschen )
            NextFreeRow( out int startRow, out int startCol );

            // calculate the Excel range to write to
            int colCount = dataArray.GetLength(0);
            int rowCount = dataArray.GetLength(1);
            if ( ( colCount == 0 ) || ( rowCount == 0 ) )
                return;   // nichts zu schreiben

            int endCol = startCol + colCount - 1;
            int endRow = startRow + rowCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber(startCol);
            string xDim2 = GetColumnFromNumber(endCol);
            string yDim1 = startRow.ToString();
            string yDim2 = endRow.ToString();

            // Excel erwartet [ zeile, spalte ], siehe ToExcelOrder
            worksheet.Range[xDim1 + yDim1, xDim2 + yDim2].Value2 = ToExcelOrder( dataArray );

        } // end: AppendRowDoublesMatrix

        /// <summary>
        /// Writes the array to the active sheet after the last row.
        /// Unter dem Datenrechteck, ab dessen erster Spalte; auf einem leeren Blatt ab A1.
        /// </summary>
        /// <param name="dataArray">double array ( eine Zeile )</param>
        public void AppendRowDoublesField(double[] dataArray)
        {
            // die Daten haben zu passen: nur endliche Zahlen ( sonst Fehlermeldung statt #ZAHL! )
            CheckFinite( dataArray, "AppendRowDoublesField" );

            // erste freie Zeile unter dem Datenrechteck ( 27.11.26: vorher SpecialCells( xlCellTypeLastCell ),
            // das merkt sich die frühere letzte Zeile auch nach dem Löschen )
            NextFreeRow( out int startRow, out int startCol );

            // calculate the Excel range to write to
            int colCount = dataArray.Length;
            if ( colCount == 0 )
                return;   // nichts zu schreiben

            int endCol = startCol + colCount - 1;

            // die Strings
            string xDim1 = GetColumnFromNumber(startCol);
            string xDim2 = GetColumnFromNumber(endCol);
            string yDim1 = startRow.ToString();

            // ein eindimensionales Feld schreibt Excel als eine Zeile ( keine Umsetzung nötig )
            worksheet.Range[xDim1 + yDim1, xDim2 + yDim1].Value2 = dataArray;

        } // end: AppendRowDoublesField

        /// <summary>
        /// Intern 2d-array of doubles will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// Die neue Mappe wird unter diesem Namen gespeichert und geschlossen; die geöffnete Mappe bleibt
        /// unverändert und wieder aktiv. Eine vorhandene Zieldatei wird nur nach Rückfrage ersetzt.
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Name des Blatts in der neuen Mappe ( leer = Excels Name )</param>
        /// <returns>true = geschrieben, false = Dialog abgebrochen, nichts geschrieben</returns>
        public bool WriteDoublesToNewTarget( string newFileTarget = "", string newTableName = "newDoubles" )
        {
            // die Daten haben zu passen: nur endliche Zahlen ( sonst Fehlermeldung statt #ZAHL! )
            CheckFinite( valuesDoubleArray, "WriteDoublesToNewTarget" );

            bool overwrite = false;
            while ( !overwrite )
            {
                // ohne Namen per Dialog; abgebrochen: nichts schreiben
                if ( !ChooseTarget( ref newFileTarget, newFileTarget != "" ) )
                    return ( false );
                // Endung und vollen Pfad ergänzen, damit File.Exists dieselbe Datei prüft, die Excel speichert
                newFileTarget = CompleteTargetName( newFileTarget );

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

            /* in die neue Mappe schreiben und sie unter dem Zielnamen speichern
             *  - 27.11.26: vorher wurde in das Blatt der geöffneten Mappe geschrieben und diese gespeichert;
             *    die Zieldatei entstand nicht, die neue Mappe blieb ungespeichert offen ( Excel hing beim Beenden )
            */
            Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
            try
            {   // neue Mappe füllen und speichern
                Excel.Worksheet newSheet = (Excel.Worksheet)newWorkbook.Worksheets[ 1 ];
                if ( !string.IsNullOrWhiteSpace( newTableName ) )
                    newSheet.Name = newTableName;

                // calculate the Excel range to write to
                int colCount = valuesDoubleArray.GetLength( 0 );
                int rowCount = valuesDoubleArray.GetLength( 1 );

                if ( ( colCount > 0 ) && ( rowCount > 0 ) )
                {   // nur bei Daten schreiben
                    int startCol = 1;
                    int endCol = startCol + colCount - 1;
                    // bei useHeader zuerst die Kopfzeile genau dieser Zahlen, die Daten darunter ( 27.11.26 )
                    int startRow = WriteHeaderRow( newSheet, colCount, headersDouble );
                    int endRow = startRow + rowCount - 1;

                    // die Strings
                    string xDim1 = GetColumnFromNumber( startCol );
                    string xDim2 = GetColumnFromNumber( endCol );
                    string yDim1 = startRow.ToString();
                    string yDim2 = endRow.ToString();

                    // Excel erwartet [ zeile, spalte ], siehe ToExcelOrder
                    newSheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = ToExcelOrder( valuesDoubleArray );
                }   // ende: nur bei Daten schreiben

                newWorkbook.SaveAs( newFileTarget );
            }   // Ende: neue Mappe füllen und speichern
            finally
            {   // neue Mappe schließen ( gespeichert ist sie schon ), die geöffnete wird wieder aktiv
                newWorkbook.Close( SaveChanges: false );
                Marshal.ReleaseComObject( newWorkbook );
            }   // Ende: neue Mappe schließen
            return ( true );

        }   // end: WriteDoublesToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// Die neue Mappe wird unter diesem Namen gespeichert und geschlossen; die geöffnete Mappe bleibt
        /// unverändert und wieder aktiv. Eine vorhandene Zieldatei wird nur nach Rückfrage ersetzt.
        /// </summary>
        /// <param name="newFileTarget">Zieldatei; leer = Dialog</param>
        /// <param name="newTableName">Name des Blatts in der neuen Mappe ( leer = Excels Name )</param>
        /// <returns>true = geschrieben, false = Dialog abgebrochen, nichts geschrieben</returns>
        public bool WriteStringsToNewTarget( string newFileTarget = "", string newTableName = "newStrings" )
        {
            // die Daten haben zu passen: es muss Texte geben, vor dem Dialog prüfen ( wie WriteDoublesToNewTarget )
            CheckStrings( valuesStringArray, "WriteStringsToNewTarget" );

            bool overwrite = false;
            while ( !overwrite )
            {
                // ohne Namen per Dialog; abgebrochen: nichts schreiben
                if ( !ChooseTarget( ref newFileTarget, newFileTarget != "" ) )
                    return ( false );
                // Endung und vollen Pfad ergänzen, damit File.Exists dieselbe Datei prüft, die Excel speichert
                newFileTarget = CompleteTargetName( newFileTarget );

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

            /* in die neue Mappe schreiben und sie unter dem Zielnamen speichern
             *  - 27.11.26: vorher wurde in das Blatt der geöffneten Mappe geschrieben und diese gespeichert;
             *    die Zieldatei entstand nicht, die neue Mappe blieb ungespeichert offen ( Excel hing beim Beenden )
            */
            Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
            try
            {   // neue Mappe füllen und speichern
                Excel.Worksheet newSheet = (Excel.Worksheet)newWorkbook.Worksheets[ 1 ];
                if ( !string.IsNullOrWhiteSpace( newTableName ) )
                    newSheet.Name = newTableName;

                // calculate the Excel range to write to
                int colCount = valuesStringArray.GetLength( 0 );
                int rowCount = valuesStringArray.GetLength( 1 );

                if ( ( colCount > 0 ) && ( rowCount > 0 ) )
                {   // nur bei Daten schreiben
                    int startCol = 1;
                    int endCol = startCol + colCount - 1;
                    // bei useHeader zuerst die Kopfzeile genau dieser Texte, die Daten darunter ( 27.11.26 )
                    int startRow = WriteHeaderRow( newSheet, colCount, headersString );
                    int endRow = startRow + rowCount - 1;

                    // die Strings
                    string xDim1 = GetColumnFromNumber( startCol );
                    string xDim2 = GetColumnFromNumber( endCol );
                    string yDim1 = startRow.ToString();
                    string yDim2 = endRow.ToString();

                    // Excel erwartet [ zeile, spalte ], siehe ToExcelOrder
                    newSheet.Range[ xDim1 + yDim1, xDim2 + yDim2 ].Value2 = ToExcelOrder( valuesStringArray );
                }   // ende: nur bei Daten schreiben

                newWorkbook.SaveAs( newFileTarget );
            }   // Ende: neue Mappe füllen und speichern
            finally
            {   // neue Mappe schließen ( gespeichert ist sie schon ), die geöffnete wird wieder aktiv
                newWorkbook.Close( SaveChanges: false );
                Marshal.ReleaseComObject( newWorkbook );
            }   // Ende: neue Mappe schließen
            return ( true );

        }   // end: WriteStringsToNewTarget

        /// <summary>
        /// Schreibt bei useHeader die gelesene Kopfzeile ( 'headers' ) in Zeile 1 eines neuen Blatts, damit sie in der
        /// neuen Datei erhalten bleibt ( 27.11.26; ohne useHeader enthält 'headers' nur Spaltenbuchstaben ).
        /// </summary>
        /// <param name="sheet">das neue Blatt</param>
        /// <param name="colCount">Spaltenzahl der Daten</param>
        /// <param name="kopf">Kopfzeile genau dieser Daten ( headersDouble bzw. headersString )</param>
        /// <returns>erste Zeile für die Daten: 2 mit Kopfzeile, sonst 1</returns>
        int WriteHeaderRow( Excel.Worksheet sheet, int colCount, string[] kopf )
        {
            if ( !useHeader || ( colCount == 0 ) || ( kopf.Length != colCount ) )
                return ( 1 );
            // ein eindimensionales Feld schreibt Excel als eine Zeile ( wie AppendRowStringField )
            sheet.Range[ "A1", GetColumnFromNumber( colCount ) + "1" ].Value2 = kopf;
            return ( 2 );

        }   // Ende: int WriteHeaderRow

    }   // end: ExcelInterop

}   // end: namespace DbaseFrame
