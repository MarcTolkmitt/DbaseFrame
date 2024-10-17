using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
