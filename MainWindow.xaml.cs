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
        /// last edit: 55.11.24
        /// </summary>
        Version version = new Version( "1.0.5" );

        DbaseFrameOleDbExcel dbfExcel;
        DbaseFrameOleDbAccess dbfAccess;

        /// <summary>
        /// standard constructor
        /// </summary>
        public MainWindow( )
        {
            InitializeComponent();

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

        // ----------------------------------------     Events

        /// <summary>
        /// handler function -> Window_Closing
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void Window_Closing( object sender, System.ComponentModel.CancelEventArgs e )
        {

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
        private void _mItemLoadExcel_Click( object sender, RoutedEventArgs e )
        {
            dbfExcel = new DbaseFrameOleDbExcel( "", false, false );
            Display( $"chosen file is {dbfExcel.fileName}" );

        }   // end: _mItemLoadExcel_Click

        /// <summary>
        /// handler function -> _mItemReadTables_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemReadTables_Click( object sender, RoutedEventArgs e )
        {
            int result = dbfExcel.ReadTableNames();
            if ( result == -1 )
            {
                Display( "no tabel found or chosen, please try again!" );
                return;

            }
            Display( $"Chosen table is number { result }" );
            Display( $"Chosen table is {dbfExcel.sheets[ result ]}" );

        }   // end: _mItemReadTables_Click

        /// <summary>
        /// handler function -> _mItemReadTableNumber_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemReadTableNumber_Click( object sender, RoutedEventArgs e )
        {
            for ( int i = 0; i < 10; i++ )
            {
                string result = dbfExcel.GetTableName( i );
                Display( $"Table {i}: {result}" );

            }

        }   // end: _mItemReadTableNumber_Click

        /// <summary>
        /// handler function -> _mItemExcelListTypesArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListTypesArray_Click( object sender, RoutedEventArgs e )
        {
            dbfExcel.ReadTypesList();
            foreach ( string[] line in dbfExcel.valuesTypes )
                Display( ArrayToString( line ) );
            Display( "\n---------------------------------" );


        }   // end: _mItemExcelListTypesArray_Click

        /// <summary>
        /// handler function -> _mItemExcelListStringArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListStringArray_Click( object sender, RoutedEventArgs e )
        {
            dbfExcel.ReadStringList();
            Display( "\nthe read data:" );
            foreach ( string[] line in dbfExcel.valuesString )
                Display( ArrayToString( line ) );
            Display( "\nstring list as jagged array:" );
            var rowArray = dbfExcel.valuesString.ToArray();
            Display( ArrayJaggedToString( rowArray, false ) );
            var listRows = rowArray.ToList();
            Display( "jagged array as list again:" );
            foreach ( string[] line in listRows )
                Display( ArrayToString( line ) );
            Display( "\n---------------------------------" );

        }   // end: _mItemExcelListStringArray_Click

        /// <summary>
        /// handler function -> _mItemExcelListDoubleArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListDoubleArray_Click( object sender, RoutedEventArgs e )
        {
            dbfExcel.ReadDoubleList();
            Display( "\nthe read data:" );
            foreach ( double[] line in dbfExcel.valuesDouble )
                Display( ArrayToString( line ) );
            Display( "\ndouble list as jagged array:" );
            var rowArray = dbfExcel.valuesDouble.ToArray();
            Display( ArrayJaggedToString( rowArray, false ) );
            var listRows = rowArray.ToList();
            Display( "jagged array as list again:" );
            foreach ( double[] line in listRows )
                Display( ArrayToString( line ) );
            Display( "\n---------------------------------" );

        }   // end: _mItemExcelListDoubleArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListDoubleToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListDoubleToExcel_Click( object sender, RoutedEventArgs e )
        {
            dbfExcel.WriteListDoubleToNewTarget();

        }   // end: _mItemWriteListDoubleToExcel_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListStringToExcel_Click( object sender, RoutedEventArgs e )
        {
            dbfExcel.WriteListStringToNewTarget();

        }   // end: _mItemWriteListStringToExcel_Click

        // --------------------------------------------------------------   ACCESS  OleDb

        /// <summary>
        /// handler function -> _mItemAccesLoad_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccesLoad_Click( object sender, RoutedEventArgs e )
        {
            dbfAccess = new DbaseFrameOleDbAccess( "", false );
            Display( $"chosen file is {dbfAccess.sourceConnectionFile}" );
            Display( $"connection string is {dbfAccess.sourceConnectionString}" );
        }   // end: _mItemAccesLoad_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessReadTables_Click( object sender, RoutedEventArgs e )
        {
            int result = dbfAccess.ReadTableNames();
            if ( result == -1 )
            {
                Display( "no tabel found or chosen, please try again!" );
                return;

            }
            Display( $"Chosen table is number {result}" );
            Display( $"Chosen table is {dbfAccess.sheets[ result ]}" );

        }

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessReadTableNumber_Click( object sender, RoutedEventArgs e )
        {
            for ( int i = 0; i < 20; i++ )
            {
                string result = dbfAccess.GetTableName( i );
                Display( $"Table {i}: {result}" );

            }

        }

        /// <summary>
        /// handler function -> _mItemAccessListTypesArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessListTypesArray_Click( object sender, RoutedEventArgs e )
        {
            dbfAccess.ReadTypesList();
            foreach ( string[] line in dbfAccess.valuesTypes )
                Display( ArrayToString( line ) );
            Display( "\n---------------------------------" );

        }   // end: _mItemAccessListTypesArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessListStringArray_Click( object sender, RoutedEventArgs e )
        {
            dbfAccess.ReadStringList();
            Display( "\nthe read data:" );
            foreach ( string[] line in dbfAccess.valuesString )
                Display( ArrayToString( line ) );
            Display( "\nstring list as jagged array:" );
            var rowArray = dbfAccess.valuesString.ToArray();
            Display( ArrayJaggedToString( rowArray, false ) );
            var listRows = rowArray.ToList();
            Display( "jagged array as list again:" );
            foreach ( string[] line in listRows )
                Display( ArrayToString( line ) );
            Display( "\n---------------------------------" );

        }   // end: _mItemAccessListStringArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemAccessListDoubleArray_Click( object sender, RoutedEventArgs e )
        {
            dbfAccess.ReadDoubleList();
            Display( "\nthe read data:" );
            foreach ( double[] line in dbfAccess.valuesDouble )
                Display( ArrayToString( line ) );
            Display( "\ndouble list as jagged array:" );
            var rowArray = dbfAccess.valuesDouble.ToArray();
            Display( ArrayJaggedToString( rowArray, false ) );
            var listRows = rowArray.ToList();
            Display( "jagged array as list again:" );
            foreach ( double[] line in listRows )
                Display( ArrayToString( line ) );
            Display( "\n---------------------------------" );

        }   // end: _mItemAccessListDoubleArray_Click

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListDoubleToAccess_Click( object sender, RoutedEventArgs e )
        {

        }

        /// <summary>
        /// handler function -> _mItemWriteListStringToExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemWriteListStringToAccess_Click( object sender, RoutedEventArgs e )
        {

        }

    }   // end: class MainWindow

}   // end: namespace DbaseFrame
