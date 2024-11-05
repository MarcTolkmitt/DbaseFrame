using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;

namespace DbaseFrame
{
    public class DbaseFrameAccess
    {
        /// <summary>
        /// created on: 05.11.24
        /// last edit: 05.11.24
        /// </summary>
        Version version = new Version( "1.0.1" );

        string sourceConnectionStart = "\"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        string sourceConnectionFile = "";
        string sourceConnectionOptions = "";
        string sourceConnectionString = "";

        public List<string[]> valuesString = new List<string[]>();
        public List<double[]> valuesDouble = new List<double[]>();

        /// <summary>
        /// Constructor for the class.
        /// </summary>
        /// <param name="file">a file name</param>
        /// <param name="silent">query for the name via dialog ?</param>
        public DbaseFrameAccess( string file = "", bool silent = true )
        {
            sourceConnectionFile = file;

            bool ok = false;
            if ( !silent )
                ok = DialogFileNameLoad( ref sourceConnectionFile );
            if ( sourceConnectionFile != "" )
            {
                sourceConnectionString =
                    sourceConnectionStart +
                    sourceConnectionFile +
                    sourceConnectionOptions;
            }
            else
                sourceConnectionString =
                    sourceConnectionStart +
                    GetDirectory() +
                    "Access_Test.accdb" +
                    sourceConnectionOptions;


        }   // end: DbaseFrameAccess ( constructor )

        public void Dummy()
        {

            // Connection string
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Path\\To\\Your\\Database.accdb";

            // Create a connection object
            OleDbConnection connection = new OleDbConnection(connectionString);

            // Open the connection
            connection.Open();

            // Execute a query to retrieve data
            OleDbCommand command = new OleDbCommand("SELECT * FROM YourTable", connection);
            OleDbDataReader reader = command.ExecuteReader();

            // Process the data

            // Close the connection
            connection.Close();
        }

        public void TableDummy()
        {

            // Replace with your Access database file path and table name
            string tableName = "Table1";

            try
            {
                using ( OleDbConnection conn = new OleDbConnection( sourceConnectionString ) )
                {
                    conn.Open();

                    OleDbCommand cmd = new OleDbCommand($"SELECT * FROM {tableName}", conn);
                    OleDbDataReader reader = cmd.ExecuteReader();

                    while ( reader.Read() )
                    {
                        // Access column values
                        int id = reader.GetInt32(0); // Replace with the actual column index or name
                        string description = reader.GetString(1);

                        Console.WriteLine( $"ID: {id}, Description: {description}" );
                    }

                    reader.Close();
                }
            }
            catch ( Exception ex )
            {
                Console.WriteLine( "Error: " + ex.Message );
            }

        }
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
        public bool DialogFileNameLoad( ref string fileName )
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

        }   // end: DialogFileNameLoad

        /// <summary>
        /// Queries a filename from the user with the standard dialog.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool DialogFileNameSave( ref string fileName )
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.SaveFileDialog();
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

        }   // end: DialogFileNameLoad



    }   // end: public class DbaseFrameAccess

}   // end: namespace DbaseFrame

