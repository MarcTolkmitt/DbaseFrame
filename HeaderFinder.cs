using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbaseFrame
{
    internal class HeaderFinder
    {
        public HeaderFinder( )
        {
            // Connection string
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
            + "FilePath "
            + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";

            // Connect to the Excel file
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();

            // Get the schema table
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            // Find the sheet you're interested in
            string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString(); // Assuming the first row is the one you want

            // Get the column headers
            DataTable columnHeaders = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null);
            string[] headerColumns = columnHeaders.AsEnumerable().Select(row => row["COLUMN_NAME"].ToString()).ToArray();

            // Use the header columns as needed

            // next guess
            bool hasHeader = false;
            string query = "SELECT COUNT(*) FROM [" + "worksheetName" + "$A1:1]";
            OleDbCommand cmd = new OleDbCommand(query, conn);
            int count = (int)cmd.ExecuteScalar();
            if ( count > 0 )
                hasHeader = true;
        }
    }
}
