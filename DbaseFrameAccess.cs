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

using System;
using System.Collections.Generic;
using System.Data;
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
        /// last edit: 06.11.24
        /// </summary>
        Version version = new Version( "1.0.2" );


        public string sourceConnectionStart = "\"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        public string sourceConnectionFile = "";
        public string sourceConnectionOptions = "";
        public string sourceConnectionString = "";

        public string targetConnectionStart = "\"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        public string targetConnectionFile = "";
        public string targetConnectionOptions = "";
        public string targetConnectionString = "";

        public List<string[]> valuesString = new List<string[]>();
        public List<double[]> valuesDouble = new List<double[]>();

        public string[] sheets = new string[1];
        public int sheetNumber = -1;

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

        /// <summary>
        /// Reads the table's data as anonymous array of
        /// strings into 'valuesString'.
        /// </summary>
        /// <param name="file">filename</param>
        /// <param name="silent">can use the file dialog</param>
        public void ReadStringList( )
        {
            using ( OleDbConnection conn = new OleDbConnection( sourceConnectionString ) )
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand( $"SELECT * FROM [{sheets[ sheetNumber ]}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                valuesString = new List<string[]>();

                while ( reader.Read() )
                {
                    string[] temp = new string[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        temp[ pos ] = 
                            reader[ pos ].ToString()
                            ?? string.Empty;
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
            using ( OleDbConnection conn = new OleDbConnection( sourceConnectionString ) )
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand($"SELECT * FROM [{sheets[ sheetNumber ]}]", conn);
                OleDbDataReader reader = command.ExecuteReader();
                valuesDouble = new List<double[]>();

                while ( reader.Read() )
                {
                    double[] temp = new double[ reader.FieldCount ];
                    for ( int pos = 0; pos < reader.FieldCount; pos++ )
                        if ( reader[ pos ].GetType() == typeof( double ) )
                            temp[ pos ] = 1.0 * reader.GetDouble( pos );
                    valuesDouble.Add( temp );

                }

            }   // end: using

        }   // end: ReadDoubleList

        /// <summary>
        /// Returns the number of a chosen table. A dialog will open to let you choose from
        /// the found table names.
        /// </summary>
        /// <returns>the number</returns>
        public int ReadTableNames( )
        {
            DataTable dt;
            using ( OleDbConnection conn = new OleDbConnection( sourceConnectionString ) )
            {
                conn.Open();
                dt =
                    conn.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null )
                    ?? new DataTable();

            }   // end: using

            if ( dt != null )
            {
                sheets = new string[ dt.Rows.Count ];

                for ( int i = 0; i < sheets.Length; i++ )
                {
                    sheets[ i ] = 
                        dt.Rows[ i ][ "TABLE_NAME" ].ToString()
                        ?? string.Empty;

                }

                DialogTablesChoice choice = new DialogTablesChoice( sheets );
                sheetNumber = choice.index;
                return ( sheetNumber );

            }
            return ( -1 );

        }   // end: ReadTableNames

        /// <summary>
        /// Direct query for the table name.
        /// </summary>
        /// <param name="numTable">number of the sheet</param>
        /// <returns>the name or 'string.empty'</returns>
        public string GetTableName( int numTable )
        {
            DataTable dt;
            using ( OleDbConnection conn = new OleDbConnection( sourceConnectionString ) )
            {
                conn.Open();
                dt =
                    conn.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null )
                    ?? new DataTable();

            }   // end: using

            if (  dt.Rows.Count > numTable )
            {
                return ( dt.Rows[ numTable ][ "TABLE_NAME" ].ToString() ?? string.Empty );
            }
            return ( string.Empty );

        }   // end: GetTableName

        public void ColumnNameDummy()
        {
            using ( OleDbConnection conn = new OleDbConnection( sourceConnectionString ) )
            {
                conn.Open();
                
                // Get the schema table for columns
                
                DataTable schemaTable = 
                    conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null )
                    ?? new DataTable();

                // Iterate through the schema table rows to access column headers (field names)
                foreach ( DataColumn column in schemaTable.Columns )
                {
                    string columnName = column.ColumnName;
                    // Use the column name as needed
                }
            }
        }

        /// <summary>
        /// Target file name for the writing is chosen. Produces the
        /// 'targetConnectionString' for convenience.
        /// </summary>
        /// <param name="file">already known ?</param>
        /// <param name="silent">use the dialog ?</param>
        public void ChooseTarget( ref string file, bool silent = true )
        {
            targetConnectionFile = file;

            bool ok = false;
            if ( !silent )
                ok = DialogFileNameSave( ref targetConnectionFile );
            if ( targetConnectionFile != "" )
            {
                targetConnectionString =
                    targetConnectionStart +
                    targetConnectionFile +
                    targetConnectionOptions;
                file = targetConnectionFile;
            }
            else
            {
                targetConnectionString =
                    sourceConnectionString +
                    GetDirectory() +
                    "newAccess_Test.accdb" +
                    targetConnectionOptions;
                file = GetDirectory() + "newAccess_Test.accdb";
            }
            targetConnectionFile = file;
            //Message.Show( file );

        }   // end: ChooseTarget


        /// <summary>
        /// Intern data list double will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        public void WriteListDoubleToNewTarget( string newFileTarget = "", string newTableName = "newDoubles" )
        {
            if ( valuesDouble.Count < 1 )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return;
            }

            bool overwrite = false;
            while ( !overwrite )
            {
                if ( newFileTarget == "" )
                    ChooseTarget( ref newFileTarget, false );
                else
                    ChooseTarget( ref newFileTarget, true );

                if ( File.Exists( newFileTarget ) )
                {
                    overwrite = Message.Ask( "Do you want to delete the file and its contents ?" );
                    if ( overwrite )
                        File.Delete( newFileTarget );

                }
                else
                    overwrite = true;
                if ( !overwrite )
                    newFileTarget = "";

            }
            // craft the 'CREATE TABLE' and 'INSERT INTO'
            int columns =  valuesDouble[0].Length;
            string tableCreateColumns = "( ";
            string tableInsertColumns = "( ";
            switch ( columns )
            {
                case 0:
                    // no data to write
                    Message.Show( "No data to write, abort!" );
                    return;
                case 1:
                    tableCreateColumns += $"{0} DOUBLE ) ";
                    tableInsertColumns += $"{0} ) VALUES ( @0 ); ";
                    break;
                case 2:
                    tableCreateColumns += $"{0} DOUBLE, ";
                    tableCreateColumns += $"{1} DOUBLE ) ";
                    tableInsertColumns += $"{0}, {1} ) VALUES ( @0, @1 ); ";
                    break;
                default:
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableCreateColumns += $"{i} DOUBLE, ";
                    tableCreateColumns += $"{( columns - 1 )} DOUBLE );";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"{i}, ";
                    tableInsertColumns += $"{( columns - 1 )} ) VALUES ( ";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"@{i}, ";
                    tableInsertColumns += $"@{( columns - 1 )} );";
                    break;

            }
            string commandCreate = $"CREATE TABLE [{newTableName}] "
                    + tableCreateColumns;
            //Message.Show( commandCreate );
            string commandInsert = $"INSERT INTO [{newTableName}] "
                    + tableInsertColumns;
            //Message.Show( commandInsert );
            using ( OleDbConnection connection = new OleDbConnection( targetConnectionString ) )
            {
                connection.Open();
                // create the table
                OleDbCommand command = new OleDbCommand( commandCreate, connection );
                command.ExecuteNonQuery();
                connection.Close();

                connection.Open();
                foreach ( double[] row in valuesDouble )
                {
                    command.CommandText = commandInsert;
                    command.Parameters.Clear();

                    for ( int pos = 0; pos < row.Length; pos++ )
                        command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }

        }   // end: WriteListDoubleToNewTarget

        /// <summary>
        /// Intern data list string will be written into a new 
        /// Excel file. If not given a name a dialog will query for it.
        /// </summary>
        /// <param name="newFileTarget"></param>
        public void WriteListStringToNewTarget( string newFileTarget = "", string newTableName = "newStrings" )
        {
            if ( valuesString.Count < 1 )
            {   // no data to write
                Message.Show( "No data to write, abort!" );
                return;
            }

            bool overwrite = false;
            while ( !overwrite )
            {
                if ( newFileTarget == "" )
                    ChooseTarget( ref newFileTarget, false );
                else
                    ChooseTarget( ref newFileTarget, true );

                if ( File.Exists( newFileTarget ) )
                {
                    overwrite = Message.Ask( "Do you want to delete the file and its contents ?" );
                    if ( overwrite )
                        File.Delete( newFileTarget );

                }
                else
                    overwrite = true;
                if ( !overwrite )
                    newFileTarget = "";

            }
            // craft the 'CREATE TABLE' and 'INSERT INTO'
            int columns =  valuesDouble[0].Length;
            string tableCreateColumns = "( ";
            string tableInsertColumns = "( ";
            switch ( columns )
            {
                case 0:
                    // no data to write
                    Message.Show( "No data to write, abort!" );
                    return;
                case 1:
                    tableCreateColumns += $"{0} VARCHAR ) ";
                    tableInsertColumns += $"{0} ) VALUES ( @0 ); ";
                    break;
                case 2:
                    tableCreateColumns += $"{0} VARCHAR, ";
                    tableCreateColumns += $"{1} VARCHAR ) ";
                    tableInsertColumns += $"{0}, {1} ) VALUES ( @0, @1 ); ";
                    break;
                default:
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableCreateColumns += $"{i} VARCHAR, ";
                    tableCreateColumns += $"{( columns - 1 )} VARCHAR );";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"{i}, ";
                    tableInsertColumns += $"{( columns - 1 )} ) VALUES ( ";
                    for ( int i = 0; i < ( columns - 1 ); i++ )
                        tableInsertColumns += $"@{i}, ";
                    tableInsertColumns += $"@{( columns - 1 )} );";
                    break;

            }
            string commandCreate = $"CREATE TABLE [{newTableName}] "
                    + tableCreateColumns;
            //Message.Show( commandCreate );
            string commandInsert = $"INSERT INTO [{newTableName}] "
                    + tableInsertColumns;
            //Message.Show( commandInsert );
            using ( OleDbConnection connection = new OleDbConnection( targetConnectionString ) )
            {
                connection.Open();
                // create the table
                OleDbCommand command = new OleDbCommand( commandCreate, connection );
                command.ExecuteNonQuery();
                connection.Close();

                connection.Open();
                foreach ( string[] row in valuesString )
                {
                    command.CommandText = commandInsert;
                    command.Parameters.Clear();

                    for ( int pos = 0; pos < row.Length; pos++ )
                        command.Parameters.AddWithValue( $"@{pos}", row[ pos ] );

                    command.ExecuteNonQuery();
                }

                connection.Close();
            }

        }   // end: WriteListStringToNewTarget

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
            dialog.DefaultExt = ".accdb"; // Default file extension
            dialog.Filter = "Access save file (.accdb)|*.accdb"; // Filter files by extension
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
            fileName = string.Empty;
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
            dialog.DefaultExt = ".accdb"; // Default file extension
            dialog.Filter = "Access save file (.accdb)|*.accdb"; // Filter files by extension
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

