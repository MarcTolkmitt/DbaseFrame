using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Controls;
using System.IO;

namespace DbaseFrame
{
    public class ExcelReadStringList
    {

        // Connect to the Excel file
        string conStringStart =
            "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=";
        string conStringEnd =
            ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";
        string connectionString;
        public List<string[]> values;

        public ExcelReadStringList( string file = "" )
        {
            if ( file != "" )
                connectionString =
                    conStringStart + file + conStringEnd;
            else
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source=" +
                    "C:\\Users\\Marc Tolkmitt\\OneDrive\\C# local für GitHub\\DbaseFrame\\" + 
                    "Parable_Demo.xlsx" +
                    ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";

            using ( OleDbConnection conn = new OleDbConnection( connectionString ) )
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand("SELECT * FROM [table0$]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                values = new List<string[]>();

                while ( reader.Read() )
                {
                    string[] temp = new string[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        temp[ pos ] = reader[ pos ].ToString();
                    values.Add( temp );

                }
                
            }

        }   // end: OleDBReadString ( constructor )

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

    }   // end: OleDBReadString

}   // end: namespace DbaseFrame
