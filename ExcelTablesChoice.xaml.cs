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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace DbaseFrame
{
    /// <summary>
    /// Interactions logic for ExcelTablesChoice.xaml
    /// </summary>
    public partial class ExcelTablesChoice : Window
    {
        /// <summary>
        /// created on: 22.01.24
        /// last edit: 29.10.24
        /// </summary>
        Version version = new Version( "1.0.2" );
        /// <summary>
        /// Chosen table's number
        /// </summary>
        public int index = -1;
        
        public ExcelTablesChoice( string[] dataStrings )
        {
            InitializeComponent();
            _listBox.ItemsSource = dataStrings.ToList();
            ShowDialog();

        }   // end: ExcelTablesChoice

        /// <summary>
        /// handler function -> _button_Click
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void _button_Click( object sender, RoutedEventArgs e )
        {
            index = _listBox.SelectedIndex;
            Close();

        }   // end: _button_Click

    }   // end: public partial class ExcelTablesChoice

}   // end: namespace DbaseFrame

