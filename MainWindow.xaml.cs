﻿/* ====================================================================
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

using System.Windows;

namespace DbaseFrame
{
    /// <summary>
    /// interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// created on: 22.01.24
        /// last edit: 15.10.24
        /// </summary>
        Version version = new Version( "1.0.3" );
        DbaseFrameExcel readExcel;

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

        /// <summary>
        /// handler function -> _mItemLoadExcel_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemLoadExcel_Click( object sender, RoutedEventArgs e )
        {
            readExcel = new DbaseFrameExcel( "", false );

        }   // end: _mItemLoadExcel_Click

        /// <summary>
        /// handler function -> _mItemReadTables_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemReadTables_Click( object sender, RoutedEventArgs e )
        {
            int result = readExcel.ReadTableNames();
            Display( $"Chosen table is number { result }" );
            Display( $"Chosen table is {readExcel.sheets[ result ]}" );

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
                string result = readExcel.GetTableName( i );
                Display( $"Table {i}: {result}" );

            }

        }   // end: _mItemReadTableNumber_Click

        /// <summary>
        /// handler function -> _mItemExcelListStringArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListStringArray_Click( object sender, RoutedEventArgs e )
        {
            readExcel.ReadStringList();
            foreach ( string[] line in readExcel.valuesString )
                Display( ArrayToString( line ) );

            var rowArray = readExcel.valuesString.ToArray();
            Display( ArrayJaggedToString( rowArray, true ) );

            var listRows = rowArray.ToList();
            foreach ( string[] line in listRows )
                Display( ArrayToString( line ) );

        }   // end: _mItemExcelListStringArray_Click

        /// <summary>
        /// handler function -> _mItemExcelListDoubleArray_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _mItemExcelListDoubleArray_Click( object sender, RoutedEventArgs e )
        {
            readExcel.ReadDoubleList();
            foreach ( double[] line in readExcel.valuesDouble )
                Display( ArrayToString( line ) );

            var rowArray = readExcel.valuesString.ToArray();
            Display( ArrayJaggedToString( rowArray, true ) );



        }   // end: _mItemExcelListDoubleArray_Click

    }   // end: class MainWindow

}   // end: namespace DbaseFrame
