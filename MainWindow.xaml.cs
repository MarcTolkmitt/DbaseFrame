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


// Ignore Spelling: Fwith

using System.Data.OleDb;
using System.Windows;

namespace DbaseFrame
{
    /// <summary>
    /// interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// created on: 22.10.24
        /// last edit: 28.11.26
        /// 27.11.26: Menü "Interop Excel" führt ExcelInterop vor; Excel- und Access-Menüpunkte über SafeDo:
        ///           Fehler ( z. B. defekte Datei ) werden angezeigt statt das Programm abzubrechen;
        ///           beim Start Paketnummer ( PackageInfo ) und Version des Beispielprogramms anzeigen;
        ///           OleDb-Menüpunkte über ExcelDo / AccessDo: ohne geöffnete Datei ein Hinweis
        ///           ( vorher "Object reference not set" ); OleDb-Excel-Dialog abgebrochen → nichts offen ( ExcelLoad );
        ///           OleDb-Schreiben zeigt das Ergebnis an ( geschrieben / nichts geschrieben );
        ///           Access-Dialog abgebrochen → nichts offen; Access-Schreibmenüs gefüllt ( vorher leer )
        /// 28.11.26: Punkt l: fehlende Daten meldet die Bibliothek als Ausnahme ( Anzeige über SafeDo ), false heißt
        ///           beim Schreiben nur noch „abgebrochen“; Menüpunkte „column names of every table / sheet“ in allen
        ///           drei Menüs ( ReadColumnNames, ShowColumnNames ); Schreiben zeigt den echten Namen ( writtenTableName )
        /// </summary>
        Version version = new Version( "1.0.7" );

        // Excel bzw. Access über OleDb; null, solange keine Datei geöffnet ist ( ExcelDo, AccessDo )
        ExcelOleDb? dbfExcel = null;
        AccessOleDb? dbfAccess = null;

        // Excel über Interop; null, solange keine Datei geöffnet ist
        ExcelInterop? dbfInterop = null;

        /// <summary>
        /// standard constructor
        /// </summary>
        public MainWindow( )
        {
            InitializeComponent();

            Display( $"{PackageInfo.Text}, Beispielprogramm {version}" );
            Display( "Init ... OK" );

        }   // end: public MainWindow



        // ---------------------------------------------     helper functions

        /// <summary>
        /// helper function, writing array data into a string
        /// </summary>
        /// <param name="data">2d ragged array </param>
        /// <returns>the data as string</returns>
        public string ArrayJaggedToString( string[][] data, bool textWrap = false )
        {
            string text = "";

            foreach ( string[] dat in data )
            {
                text += $" [ {string.Join( ", ", dat )} ] ";
                if ( textWrap )
                    text += "\n";

            }
            text += "\n";
            return ( text );

        }   // end: ArrayToString

        /// <summary>
        /// helper function, writing array data into a string
        /// </summary>
        /// <param name="data">2d ragged array </param>
        /// <returns>the data as string</returns>
        public string ArrayJaggedToString( double[][] data, bool textWrap = false )
        {
            string text = "";

            foreach ( double[] dat in data )
            {
                text += $" [ {string.Join( ", ", dat )} ] ";
                if ( textWrap )
                    text += "\n";

            }
            text += "\n";
            return ( text );

        }   // end: ArrayToString

        /// <summary>
        /// helper function, writing array data into a string
        /// </summary>
        /// <param name="data">array </param>
        /// <returns>the data as string</returns>
        public string ArrayToString( string[] data )
        {
            string text = "";

            foreach ( var dat in data )
            {
                text += $" [ {string.Join( ", ", dat )} ] ";

            }
            //text += "\n";
            return ( text );

        }   // end: ArrayToString

        /// <summary>
        /// helper function, writing array data into a string
        /// </summary>
        /// <param name="data">array </param>
        /// <returns>the data as string</returns>
        public string ArrayToString( double[] data )
        {
            string text = "";

            foreach ( var dat in data )
            {
                text += $" [ {string.Join( ", ", dat )} ] ";

            }
            //text += "\n";
            return ( text );

        }   // end: ArrayToString

        /// <summary>
        /// helper function to write the text into the main window
        /// </summary>
        /// <param name="text">input string</param>
        public void Display( string? text )
        {
            if ( !string.IsNullOrEmpty( text ) )
                textBlock.Text += text + "\n";
            textScroll.ScrollToBottom();

        }   // end: Display

        /// <summary>
        /// helper function to write the text into the main window
        /// </summary>
        /// <param name="text">any-object-variant</param>
        private void Display( Object obj )
        {
            Display( obj.ToString() );

        }   // end: Display

        /// <summary>
        /// Hilfsfunktion: führt einen Menüpunkt aus und zeigt Fehler an, statt das Programm abbrechen zu lassen
        /// ( z. B. defekte oder fehlende Datei, noch keine Datei bzw. Tabelle gewählt ).
        /// </summary>
        /// <param name="quelle">Name für die Meldung ( Excel, Access )</param>
        /// <param name="aktion">der Menüpunkt</param>
        void SafeDo( string quelle, Action aktion )
        {
            try
            {
                aktion( );
            }
            catch ( Exception ex )
            {   // Fehler anzeigen, Programm läuft weiter
                Display( $"{quelle}-Fehler: {ex.Message}" );
            }   // Ende: Fehler anzeigen, Programm läuft weiter

        }   // Ende: void SafeDo

        /// <summary>
        /// Hilfsfunktion für die drei Menüpunkte ReadColumnNames ( 28.11.26 ): zeigt für jede Tabelle bzw. jedes Blatt die
        /// Spaltennamen. Ein Fehler einer Tabelle ( z. B. leeres Blatt ) steht in ihrer Zeile, die übrigen folgen trotzdem.
        /// </summary>
        /// <param name="tabelle">Name der Tabelle Nr. i, leer = keine weitere ( GetTableName / GetSheetName )</param>
        /// <param name="spalten">Spaltennamen der Tabelle Nr. i ( ReadColumnNames )</param>
        void ShowColumnNames( Func<int, string> tabelle, Func<int, string[]> spalten )
        {
            int nummer = 0;
            for ( string name = tabelle( 0 ); name != ""; name = tabelle( ++nummer ) )
            {   // je Tabelle
                string text;
                try
                {   // Spaltennamen
                    text = string.Join( " | ", spalten( nummer ) );
                }   // Ende: Spaltennamen
                catch ( Exception ex )
                {   // Fehler dieser Tabelle anzeigen, weiter mit der nächsten
                    text = "Fehler: " + ex.Message;
                }   // Ende: Fehler dieser Tabelle anzeigen
                Display( $"Tabelle {nummer} '{name}': {text}" );
            }   // Ende: je Tabelle
            if ( nummer == 0 )
                Display( "keine Tabelle gefunden" );

        }   // Ende: void ShowColumnNames

        /// <summary>
        /// Hilfsfunktion: führt einen Menüpunkt mit der geöffneten Excel-Datei ( OleDb ) aus und zeigt Fehler an
        /// ( wie InteropDo ); ohne geöffnete Datei ein Hinweis statt "Object reference not set".
        /// </summary>
        /// <param name="aktion">der Menüpunkt</param>
        void ExcelDo( Action<ExcelOleDb> aktion )
        {
            if ( dbfExcel == null )
            {   // nichts offen
                Display( "OleDb Excel: zuerst eine Datei öffnen" );
                return;
            }   // ende: nichts offen
            ExcelOleDb datei = dbfExcel;
            SafeDo( "Excel", ( ) => aktion( datei ) );

        }   // Ende: void ExcelDo

        /// <summary>
        /// Hilfsfunktion: führt einen Menüpunkt mit der geöffneten Access-Datei aus und zeigt Fehler an
        /// ( wie InteropDo ); ohne geöffnete Datei ein Hinweis statt "Object reference not set".
        /// </summary>
        /// <param name="aktion">der Menüpunkt</param>
        void AccessDo( Action<AccessOleDb> aktion )
        {
            if ( dbfAccess == null )
            {   // nichts offen
                Display( "OleDb Access: zuerst eine Datei öffnen" );
                return;
            }   // ende: nichts offen
            AccessOleDb datei = dbfAccess;
            SafeDo( "Access", ( ) => aktion( datei ) );

        }   // Ende: void AccessDo

        // ----------------------------------------     Events

        /// <summary>
        /// handler function -> Window_Closing
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void Window_Closing( object sender, System.ComponentModel.CancelEventArgs e )
        {
            // Excel beenden, falls über Interop noch eine Datei offen ist ( sonst bleibt Excel im Hintergrund )
            InteropClose( );

        }   // end: private void Window_Closing

        /// <summary>
        /// handler function -> MenuItem
        /// used for exit routines
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void MenuQuit_Click( object sender, RoutedEventArgs e )
        {
            this.Close();

        }   // end: MenuQuit_Click

        // ----------------------------------------------   EXCEL OleDb

        /// <summary>
        /// handler function -> _mItemLoadExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemLoadExcelNH_Click( object sender, RoutedEventArgs e )
        {   // use dialog for the filename and no headers
            ExcelLoad( false );

        }   // end: _mItemLoadExcelNH_Click

        /// <summary>
        /// Hilfsfunktion: Excel-Datei ( OleDb ) per Dialog wählen; abgebrochen → nichts offen ( wie InteropOpen ).
        /// </summary>
        /// <param name="headers">erste Zeile als Kopfzeile</param>
        void ExcelLoad( bool headers )
        {
            ExcelOleDb datei = new ExcelOleDb( "", false, headers );
            if ( datei.lastError != "" )
            {   // abgebrochen: nichts offen ( 27.11.26: vorher stillschweigend Parable_Demo.xlsx )
                dbfExcel = null;
                Display( "OleDb Excel: keine Datei geöffnet – " + datei.lastError );
                return;
            }   // ende: abgebrochen
            dbfExcel = datei;
            Display( $"chosen file is {datei.fileName}" );

        }   // Ende: void ExcelLoad

        /// <summary>
        /// handler function -> _mItemLoadExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemLoadExcelWH_Click( object sender, RoutedEventArgs e )
        {   // use dialog for the filename and headers
            ExcelLoad( true );

        }   // end: _mItemLoadExcelWH_Click

        /// <summary>
        /// handler function -> _mItemReadTables_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemReadTables_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x =>
            {
                int result = x.ReadTableNames();
                if ( result == -1 )
                {
                    Display( "no tabel found or chosen, please try again!" );
                    return;

                }
                Display( $"Chosen table is number { result }" );
                Display( $"Chosen table is {x.sheets[ result ]}" );
            } );

        }   // end: _mItemReadTables_Click

        /// <summary>
        /// handler function -> _mItemReadTableNumber_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemReadTableNumber_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x =>
            {
                for ( int i = 0; i < 10; i++ )
                {
                    string result = x.GetTableName( i );
                    Display( $"Table {i}: {result}" );

                }
            } );

        }   // end: _mItemReadTableNumber_Click

        /// <summary>
        /// handler function -> _mItemReadColumnNames_Click ( 28.11.26 ): Spaltennamen jeder Tabelle ( ReadColumnNames )
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemReadColumnNames_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x => ShowColumnNames( x.GetTableName, x.ReadColumnNames ) );

        }   // end: _mItemReadColumnNames_Click

        /// <summary>
        /// handler function -> _mItemExcelListTypesArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListTypesArray_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x =>
            {
                x.ReadTypesList();
                foreach ( string[] line in x.valuesTypes )
                    Display( ArrayToString( line ) );
                Display( "\n---------------------------------" );
            } );

        }   // end: _mItemExcelListTypesArray_Click

        /// <summary>
        /// handler function -> _mItemExcelListStringArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListStringArray_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x =>
            {
                x.ReadStringList();
                Display( "\nthe read data:" );
                foreach ( string[] line in x.valuesString )
                    Display( ArrayToString( line ) );
                Display( "\nstring list as jagged array:" );
                var rowArray = x.valuesString.ToArray();
                Display( ArrayJaggedToString( rowArray, false ) );
                var listRows = rowArray.ToList();
                Display( "jagged array as list again:" );
                foreach ( string[] line in listRows )
                    Display( ArrayToString( line ) );
                Display( "\n---------------------------------" );
            } );

        }   // end: _mItemExcelListStringArray_Click

        /// <summary>
        /// handler function -> _mItemExcelListDoubleArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListDoubleArray_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x =>
            {
                x.ReadDoubleList();
                Display( "\nthe read data:" );
                foreach ( double[] line in x.valuesDouble )
                    Display( ArrayToString( line ) );
                Display( "\ndouble list as jagged array:" );
                var rowArray = x.valuesDouble.ToArray();
                Display( ArrayJaggedToString( rowArray, false ) );
                var listRows = rowArray.ToList();
                Display( "jagged array as list again:" );
                foreach ( double[] line in listRows )
                    Display( ArrayToString( line ) );
                Display( "\n---------------------------------" );
            } );

        }   // end: _mItemExcelListDoubleArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListDoubleToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListDoubleToExcel_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x =>
            {
                Display( x.WriteListDoubleToNewTarget( )
                    ? $"OleDb Excel: Zahlen in eine neue Datei geschrieben ( Tabelle '{x.writtenTableName}' )"
                    : "OleDb Excel: abgebrochen, nichts geschrieben" );
            } );

        }   // end: _mItemWriteListDoubleToExcel_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListStringToExcel_Click( object sender, RoutedEventArgs e )
        {
            ExcelDo( x =>
            {
                Display( x.WriteListStringToNewTarget( )
                    ? $"OleDb Excel: Texte in eine neue Datei geschrieben ( Tabelle '{x.writtenTableName}' )"
                    : "OleDb Excel: abgebrochen, nichts geschrieben" );
            } );

        }   // end: _mItemWriteListStringToExcel_Click

        // --------------------------------------------------------------   ACCESS  OleDb

        /// <summary>
        /// handler function -> _mItemAccesLoad_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccesLoad_Click( object sender, RoutedEventArgs e )
        {
            SafeDo( "Access", ( ) =>
            {
                AccessOleDb datei = new AccessOleDb( "", false );
                if ( datei.lastError != "" )
                {   // abgebrochen: nichts offen ( 27.11.26, wie ExcelLoad )
                    dbfAccess = null;
                    Display( "OleDb Access: keine Datei geöffnet – " + datei.lastError );
                    return;
                }   // ende: abgebrochen
                dbfAccess = datei;
                Display( $"chosen file is {datei.sourceConnectionFile}" );
                Display( $"connection string is {datei.sourceConnectionString}" );
            } );

        }   // end: _mItemAccesLoad_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessReadTables_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x =>
            {
                int result = x.ReadTableNames();
                if ( result == -1 )
                {
                    Display( "no tabel found or chosen, please try again!" );
                    return;

                }
                Display( $"Chosen table is number {result}" );
                Display( $"Chosen table is {x.sheets[ result ]}" );
            } );

        }

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessReadTableNumber_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x =>
            {
                for ( int i = 0; i < 20; i++ )
                {
                    string result = x.GetTableName( i );
                    Display( $"Table {i}: {result}" );

                }
            } );

        }

        /// <summary>
        /// handler function -> _mItemAccessReadColumnNames_Click ( 28.11.26 ): Spaltennamen jeder Tabelle ( ReadColumnNames )
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessReadColumnNames_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x => ShowColumnNames( x.GetTableName, x.ReadColumnNames ) );

        }   // end: _mItemAccessReadColumnNames_Click

        /// <summary>
        /// handler function -> _mItemAccessListTypesArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessListTypesArray_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x =>
            {
                x.ReadTypesList();
                foreach ( string[] line in x.valuesTypes )
                    Display( ArrayToString( line ) );
                Display( "\n---------------------------------" );
            } );

        }   // end: _mItemAccessListTypesArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessListStringArray_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x =>
            {
                x.ReadStringList();
                Display( "\nthe read data:" );
                foreach ( string[] line in x.valuesString )
                    Display( ArrayToString( line ) );
                Display( "\nstring list as jagged array:" );
                var rowArray = x.valuesString.ToArray();
                Display( ArrayJaggedToString( rowArray, false ) );
                var listRows = rowArray.ToList();
                Display( "jagged array as list again:" );
                foreach ( string[] line in listRows )
                    Display( ArrayToString( line ) );
                Display( "\n---------------------------------" );
            } );

        }   // end: _mItemAccessListStringArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessListDoubleArray_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x =>
            {
                x.ReadDoubleList();
                Display( "\nthe read data:" );
                foreach ( double[] line in x.valuesDouble )
                    Display( ArrayToString( line ) );
                Display( "\ndouble list as jagged array:" );
                var rowArray = x.valuesDouble.ToArray();
                Display( ArrayJaggedToString( rowArray, false ) );
                var listRows = rowArray.ToList();
                Display( "jagged array as list again:" );
                foreach ( double[] line in listRows )
                    Display( ArrayToString( line ) );
                Display( "\n---------------------------------" );
            } );

        }   // end: _mItemAccessListDoubleArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListDoubleToAccess_Click ( 27.11.26: vorher leer )
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListDoubleToAccess_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x =>
            {
                Display( x.WriteListDoubleToNewTarget( )
                    ? $"OleDb Access: Zahlen in eine neue Datei geschrieben ( Tabelle '{x.writtenTableName}' )"
                    : "OleDb Access: abgebrochen, nichts geschrieben" );
            } );

        }   // end: _mItemWriteListDoubleToAccess_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToAccess_Click ( 27.11.26: vorher leer )
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListStringToAccess_Click( object sender, RoutedEventArgs e )
        {
            AccessDo( x =>
            {
                Display( x.WriteListStringToNewTarget( )
                    ? $"OleDb Access: Texte in eine neue Datei geschrieben ( Tabelle '{x.writtenTableName}' )"
                    : "OleDb Access: abgebrochen, nichts geschrieben" );
            } );

        }   // end: _mItemWriteListStringToAccess_Click

        // ----------------------------------------------   EXCEL Interop ( 27.11.26 )

        /* Vorführung von ExcelInterop
         *  - Matrizen sind Matrix( x, y ): erster Index Spalte, zweiter Zeile; angezeigt wird Zeile für Zeile
         *  - die Daten haben zu passen: Fehlermeldungen der Klasse ( z. B. leere Rand-Zeile ) werden angezeigt
         *  - Excel läuft unsichtbar im Hintergrund, bis "close" oder das Fenster geschlossen wird
        */

        /// <summary>
        /// Hilfsfunktion: Matrix( x, y ) Zeile für Zeile als Text.
        /// </summary>
        /// <param name="matrix">Matrix( x, y )</param>
        /// <returns>eine Zeile Text je Zeile der Matrix</returns>
        public string MatrixToString<T>( T[,] matrix )
        {
            string text = "";
            for ( int y = 0; y < matrix.GetLength( 1 ); y++ )
            {   // je Zeile
                var zeile = new List<string>( );
                for ( int x = 0; x < matrix.GetLength( 0 ); x++ )
                    zeile.Add( matrix[ x, y ]?.ToString( ) ?? "" );
                text += $" [ {string.Join( ", ", zeile )} ]\n";
            }   // Ende: je Zeile
            return ( text );

        }   // Ende: string MatrixToString

        /// <summary>
        /// Hilfsfunktion: Datei schließen und Excel beenden, falls offen.
        /// </summary>
        void InteropClose( )
        {
            if ( dbfInterop != null )
            {   // offene Mappe
                dbfInterop.ExcelInteropQuit( );
                dbfInterop = null;
                Display( "Interop: Datei geschlossen, Excel beendet" );
            }   // ende: offene Mappe

        }   // Ende: void InteropClose

        /// <summary>
        /// Hilfsfunktion: Datei per Dialog öffnen ( vorher geöffnete wird geschlossen ).
        /// </summary>
        /// <param name="headers">erste Zeile als Kopfzeile</param>
        /// <param name="readOnly">nur lesen</param>
        void InteropOpen( bool headers, bool readOnly )
        {
            InteropClose( );
            dbfInterop = new ExcelInterop( "", false, headers, readOnly );
            if ( dbfInterop.workbookOpen )
            {   // geöffnet; nur lesen kann Excel erzwingen ( Datei woanders geöffnet ), dann Hinweis
                Display( $"Interop: geöffnet {dbfInterop.fileName} ( Kopfzeile: {headers}, nur lesen: {dbfInterop.readOnly} )" );
                if ( dbfInterop.lastError != "" )
                    Display( "Interop: " + dbfInterop.lastError );
            }   // Ende: geöffnet
            else
            {   // abgebrochen oder Fehler: Grund anzeigen
                Display( "Interop: keine Datei geöffnet – " + dbfInterop.lastError );
                dbfInterop.ExcelInteropQuit( );
                dbfInterop = null;
            }   // ende: abgebrochen oder Fehler

        }   // Ende: void InteropOpen

        /// <summary>
        /// Hilfsfunktion: führt eine Aktion mit der offenen Datei aus und zeigt Fehlermeldungen an.
        /// </summary>
        /// <param name="aktion">die Aktion</param>
        void InteropDo( Action<ExcelInterop> aktion )
        {
            if ( dbfInterop == null )
            {   // nichts offen
                Display( "Interop: zuerst eine Datei öffnen" );
                return;
            }   // ende: nichts offen
            try
            {
                aktion( dbfInterop );
            }
            catch ( Exception ex )
            {   // z. B. die Daten passen nicht
                Display( "Interop-Fehler: " + ex.Message );
            }   // Ende: z. B. die Daten passen nicht

        }   // Ende: void InteropDo

        private void _mItemInteropOpenNH_Click( object sender, RoutedEventArgs e )
        {   // Dialog, ohne Kopfzeile, nur lesen
            InteropOpen( false, true );

        }   // end: _mItemInteropOpenNH_Click

        private void _mItemInteropOpenWH_Click( object sender, RoutedEventArgs e )
        {   // Dialog, mit Kopfzeile, nur lesen
            InteropOpen( true, true );

        }   // end: _mItemInteropOpenWH_Click

        private void _mItemInteropOpenWrite_Click( object sender, RoutedEventArgs e )
        {   // Dialog, mit Kopfzeile, beschreibbar
            InteropOpen( true, false );

        }   // end: _mItemInteropOpenWrite_Click

        private void _mItemInteropSheetNames_Click( object sender, RoutedEventArgs e )
        {   // alle Blattnamen, 0-basiert
            InteropDo( x =>
            {
                for ( int nummer = 0; x.GetSheetName( nummer ) != ""; nummer++ )
                    Display( $"Blatt {nummer}: {x.GetSheetName( nummer )}" );
            } );

        }   // end: _mItemInteropSheetNames_Click

        private void _mItemInteropColumnNames_Click( object sender, RoutedEventArgs e )
        {   // Spaltennamen jedes Blatts ( ReadColumnNames, 28.11.26 ), Blattwahl bleibt unverändert
            InteropDo( x => ShowColumnNames( x.GetSheetName, x.ReadColumnNames ) );

        }   // end: _mItemInteropColumnNames_Click

        private void _mItemInteropChooseSheet_Click( object sender, RoutedEventArgs e )
        {   // Blatt im Dialog wählen
            InteropDo( x =>
            {
                int nummer = x.ReadSheetNames( );
                Display( nummer >= 0
                    ? $"Interop: Blatt {nummer} '{x.GetSheetName( nummer )}' gewählt"
                    : "Interop: kein Blatt gewählt, es gilt das aktive Blatt" );
            } );

        }   // end: _mItemInteropChooseSheet_Click

        private void _mItemInteropTypes_Click( object sender, RoutedEventArgs e )
        {   // Typen der Zellwerte ( ganzes Rechteck, ohne Prüfung )
            InteropDo( x =>
            {
                x.ReadTypes( );
                Display( "Typen ( Zeile für Zeile ):\n" + MatrixToString( x.valuesTypesArray ) );
            } );

        }   // end: _mItemInteropTypes_Click

        private void _mItemInteropStrings_Click( object sender, RoutedEventArgs e )
        {   // Texte, Kopfzeile getrennt
            InteropDo( x =>
            {
                x.ReadStrings( );
                Display( "Kopfzeile: " + ArrayToString( x.headers ) );
                Display( "Texte ( Zeile für Zeile ):\n" + MatrixToString( x.valuesStringArray ) );
            } );

        }   // end: _mItemInteropStrings_Click

        private void _mItemInteropDoubles_Click( object sender, RoutedEventArgs e )
        {   // Zahlen, Kopfzeile getrennt
            InteropDo( x =>
            {
                x.ReadDoubles( );
                Display( "Kopfzeile: " + ArrayToString( x.headers ) );
                Display( "Zahlen ( Zeile für Zeile ):\n" + MatrixToString( x.valuesDoubleArray ) );
            } );

        }   // end: _mItemInteropDoubles_Click

        private void _mItemInteropRange_Click( object sender, RoutedEventArgs e )
        {   // fester Bereich, X = Spalte, Y = Zeile
            InteropDo( x =>
            {
                Display( "Bereich A1:C3:\n" + MatrixToString( x.ReadStrings( 1, 1, 3, 3 ) ) );
            } );

        }   // end: _mItemInteropRange_Click

        private void _mItemInteropWriteDoubles_Click( object sender, RoutedEventArgs e )
        {   // gelesene Zahlen in eine neue Datei ( Dialog )
            InteropDo( x =>
            {
                if ( x.valuesDoubleArray == null )
                {   // noch nichts gelesen
                    Display( "Interop: zuerst die Zahlen lesen" );
                    return;
                }   // ende: noch nichts gelesen
                Display( x.WriteDoublesToNewTarget( )
                    ? $"Interop: Zahlen in eine neue Datei geschrieben ( Blatt '{x.writtenTableName}' )"
                    : "Interop: abgebrochen, nichts geschrieben" );
            } );

        }   // end: _mItemInteropWriteDoubles_Click

        private void _mItemInteropWriteStrings_Click( object sender, RoutedEventArgs e )
        {   // gelesene Texte in eine neue Datei ( Dialog )
            InteropDo( x =>
            {
                if ( x.valuesStringArray == null )
                {   // noch nichts gelesen
                    Display( "Interop: zuerst die Texte lesen" );
                    return;
                }   // ende: noch nichts gelesen
                Display( x.WriteStringsToNewTarget( )
                    ? $"Interop: Texte in eine neue Datei geschrieben ( Blatt '{x.writtenTableName}' )"
                    : "Interop: abgebrochen, nichts geschrieben" );
            } );

        }   // end: _mItemInteropWriteStrings_Click

        private void _mItemInteropAppend_Click( object sender, RoutedEventArgs e )
        {   // eine Zeile unter das Datenrechteck des gewählten Blatts, dann speichern
            InteropDo( x =>
            {
                if ( x.readOnly )
                {   // nur lesen
                    Display( "Interop: Datei ist nur zum Lesen geöffnet ( 'open … for writing' wählen )" );
                    return;
                }   // ende: nur lesen
                x.AppendRowDoublesField( new double[] { 1, 2, 3 } );
                Display( x.SaveWorkbook( ) ? "Interop: Zeile angehängt und gespeichert" : "Interop: Speichern fehlgeschlagen" );
            } );

        }   // end: _mItemInteropAppend_Click

        private void _mItemInteropNewSheet_Click( object sender, RoutedEventArgs e )
        {   // neues Blatt am Ende mit einer Matrix( x, y ), dann speichern
            InteropDo( x =>
            {
                if ( x.readOnly )
                {   // nur lesen
                    Display( "Interop: Datei ist nur zum Lesen geöffnet ( 'open … for writing' wählen )" );
                    return;
                }   // ende: nur lesen
                int nummer = x.NewSheet( "Demo" );
                // Matrix( x, y ): 3 Spalten, 2 Zeilen; [ 2, 0 ] steht in C1, [ 0, 1 ] in A2
                double[,] matrix = new double[ 3, 2 ];
                for ( int spalte = 0; spalte < 3; spalte++ )
                    for ( int zeile = 0; zeile < 2; zeile++ )
                        matrix[ spalte, zeile ] = 10 * ( spalte + 1 ) + ( zeile + 1 );
                x.WriteDoubles( matrix, 1, 1 );
                Display( $"Interop: Blatt {nummer} '{x.GetSheetName( nummer )}' angelegt:\n" + MatrixToString( matrix ) );
                Display( x.SaveWorkbook( ) ? "Interop: gespeichert" : "Interop: Speichern fehlgeschlagen" );
            } );

        }   // end: _mItemInteropNewSheet_Click

        private void _mItemInteropClose_Click( object sender, RoutedEventArgs e )
        {   // Datei schließen, Excel beenden
            if ( dbfInterop == null )
                Display( "Interop: keine Datei offen" );
            InteropClose( );

        }   // end: _mItemInteropClose_Click

    }   // end: class MainWindow

}   // end: namespace DbaseFrame
