using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Xml;

namespace DbaseFrame
{
    public class ExcelWriteStringList
    {
        // Connect to the Excel file
        string conStringStart =
            "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=";
        string conStringEnd =
            ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";
        string connectionString = "";
        public List<string[]> values;

        public ExcelWriteStringList( string file = "", bool silent = true )
        {
            string fileName = file;
            bool ok = false;
            if ( !silent )
                ok = DialogFileName( ref fileName );
            if ( fileName != "" )
                connectionString =
                    conStringStart + fileName + conStringEnd;
            else
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source=" +
                    "C:\\" +
                    "Write_Excel_Demo.xlsx" +
                    ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";
            string createTable = "CREATE TABLE table_name(" +
                "column1 data_type,column2 data_type, " +
                ")CONSTRAINT constraint_name[ PRIMARY KEY | UNIQUE | FOREIGN KEY ]);";

            using ( OleDbConnection conn = new OleDbConnection( connectionString ) )
            {
                conn.Open( );
                OleDbCommand command = new OleDbCommand("INSERT INTO [Sheet1$] (Column1, Column2) VALUES (@Value1, @Value2)", conn);
                OleDbDataReader reader = command.ExecuteReader();
                values = new List<string[ ]>( );

                while ( reader.Read( ) )
                {
                    string[] temp = new string[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        temp[ pos ] = reader[ pos ].ToString( );
                    values.Add( temp );

                }

            }


        }   // end: public ExcelWriteStringList ( constructor )

        /// <summary>
        /// Creates a string filled with the numbering of the columns
        /// for the SQL query.
        /// </summary>
        /// <param name="number">the length of the array</param>
        /// <returns>the string</returns>
        public string GetColumnNumber( int number )
        {
            string temp = "(";
            for ( int i = 0; i < ( number - 1 ); i++ )
                temp += $"{i},";
            temp += $"{( number - 1 )})";
            return ( temp );

        }   // end: GetColumnNumber

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
        public bool DialogFileName( ref string fileName )
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = fileName; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel save file (.xlsx)|*.xlsx"; // Filter files by extension
            dialog.DefaultDirectory = GetDirectory( );

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

        }   // end: DialogFileName

    }   // end: public class ExcelWriteStringList

}   // end: namespace DbaseFrame
