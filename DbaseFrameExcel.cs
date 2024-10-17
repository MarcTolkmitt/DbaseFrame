using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Controls;
using System.IO;
using System.Net;
using System.Windows.Media.Animation;
using System.Configuration;
using System.Data;

namespace DbaseFrame
{
    public class DbaseFrameExcel
    {

        // Connect to the Excel file
        string conStringStart =
            "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=";
        string conStringEnd =
            ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";
        string connectionString = "";
        string fileName = "";
        public List<string[]> valuesString = new List<string[]>();
        public List<double[]> valuesDouble = new List<double[]>();

        /// <summary>
        /// Constructor for the class. 
        /// </summary>
        /// <param name="file">a file name</param>
        /// <param name="silent">query for the name via dialog ?</param>
        public DbaseFrameExcel( string file = "", bool silent = true )
        {
            fileName = file;
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
                    "Parable_Demo.xlsx" +
                    ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";


        }   // end: DbaseFrameExcel ( constructor )

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// strings into 'valuesString'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadStringList( )
        {
            using ( OleDbConnection conn = new OleDbConnection( connectionString ) )
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand("SELECT * FROM [table0$]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                valuesString = new List<string[]>();

                while ( reader.Read() )
                {
                    string[] temp = new string[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        temp[ pos ] = reader.GetString( pos );
                    valuesString.Add( temp );

                }
                
            }   // end: using

        }   // end: ReadStringList

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// doubles into 'valuesDouble'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadDoubleList( )
        {
            using ( OleDbConnection conn = new OleDbConnection( connectionString ) )
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand("SELECT * FROM [table0$]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                valuesDouble = new List<double[]>();

                while ( reader.Read() )
                {
                    double[] temp = new double[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        temp[ pos ] = reader.GetDouble( pos );
                    valuesDouble.Add( temp );

                }

            }   // end: using

        }   // end: ReadDoubleList

        public int ReadTableNames()
        {
            DataTable? dt;
            using ( OleDbConnection conn = new OleDbConnection( connectionString ) )
            {
                conn.Open();
                dt = 
                    conn.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );
            }   // end: using
            if ( dt != null )
            {
                string[] sheets = new string[ dt.Rows.Count ];
                int i = 0;
                foreach( DataRow row in dt.Rows )
                {
                    sheets[ i ] = row[ "TABLE_NAME" ].ToString();
                    i++;

                }
                ExcelTablesChoice choice = new ExcelTablesChoice( sheets, ref i );
                return( i );

            }
            return( -1  );

        }   // end: ReadTableNames

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
        public bool DialogFileName( ref string fileName )
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

        }   // end: DialogFileName

    }   // end: OleDBReadString

}   // end: namespace DbaseFrame
