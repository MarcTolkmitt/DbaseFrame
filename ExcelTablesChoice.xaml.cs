// Ignore Spelling: Dbase

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
        public string[] items = new string[ 3 ] { "Hund", "Katze", "Maus" };
        int index = -1;
        
        public ExcelTablesChoice( string[] dataStrings, ref int chosenIndex )
        {
            InitializeComponent();
            index = chosenIndex;
            _listBox.ItemsSource = dataStrings.ToList();
            Show();

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

