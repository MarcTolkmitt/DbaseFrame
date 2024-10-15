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
        /// <summary>
        /// standard constructor
        /// </summary>
        public MainWindow( )
        {
            InitializeComponent();

            Display( "Init ... ok" );

        }   // end: public MainWindow

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



        // ---------------------------------------------     helper functions

        /// <summary>
        /// helper function, writing array data into a string
        /// </summary>
        /// <param name="data">2d ragged array </param>
        /// <returns>the data as string</returns>
        public string ArrayToString( double[][] data, bool textWrap = false )
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
        private void Display( int obj )
        {
            Display( obj.ToString() );

        }   // end: Display

    }   // end: class MainWindow

}   // end: namespace DbaseFrame
