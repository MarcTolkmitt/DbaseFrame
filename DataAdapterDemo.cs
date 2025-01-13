using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Odbc;

namespace DbaseFrame
{
    class DataAdapterDemo
    {
        public static OleDbDataAdapter CreateDataAdapter(
            OleDbConnection connection )
        {
            string selectCommand =
        "SELECT CustomerID, CompanyName FROM Customers";
            OleDbDataAdapter adapter =
        new OleDbDataAdapter(selectCommand, connection);

            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;

            // Create the Insert, Update and Delete commands.
            adapter.InsertCommand = new OleDbCommand(
                "INSERT INTO Customers (CustomerID, CompanyName) " +
                "VALUES (?, ?)" );

            adapter.UpdateCommand = new OleDbCommand(
                "UPDATE Customers SET CustomerID = ?, CompanyName = ? " +
                "WHERE CustomerID = ?" );

            adapter.DeleteCommand = new OleDbCommand(
                "DELETE FROM Customers WHERE CustomerID = ?" );

            // Create the parameters.
            adapter.InsertCommand.Parameters.Add( "@CustomerID",
                OleDbType.Char, 5, "CustomerID" );
            adapter.InsertCommand.Parameters.Add( "@CompanyName",
                OleDbType.VarChar, 40, "CompanyName" );

            adapter.UpdateCommand.Parameters.Add( "@CustomerID",
                OleDbType.Char, 5, "CustomerID" );
            adapter.UpdateCommand.Parameters.Add( "@CompanyName",
                OleDbType.VarChar, 40, "CompanyName" );
            adapter.UpdateCommand.Parameters.Add( "@oldCustomerID",
                OleDbType.Char, 5, "CustomerID" ).SourceVersion =
                DataRowVersion.Original;

            adapter.DeleteCommand.Parameters.Add( "@CustomerID",
                OleDbType.Char, 5, "CustomerID" ).SourceVersion =
                DataRowVersion.Original;

            return adapter;
        }

        public static OleDbDataAdapter CreateCustomerAdapter(
            OleDbConnection connection )
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            OleDbCommand command;

            // Create the SelectCommand.
            command = new OleDbCommand( "SELECT CustomerID FROM Customers " +
                "WHERE Country = ? AND City = ?", connection );

            command.Parameters.Add( "Country", OleDbType.VarChar, 15 );
            command.Parameters.Add( "City", OleDbType.VarChar, 15 );

            adapter.SelectCommand = command;

            // Create the InsertCommand.
            command = new OleDbCommand(
                "INSERT INTO Customers (CustomerID, CompanyName) " +
                "VALUES (?, ?)", connection );

            command.Parameters.Add(
                "CustomerID", OleDbType.Char, 5, "CustomerID" );
            command.Parameters.Add(
                "CompanyName", OleDbType.VarChar, 40, "CompanyName" );

            adapter.InsertCommand = command;
            return adapter;
        }

        public static OleDbDataAdapter CreateCustomerAdapter(
            OleDbConnection connection )
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            OleDbCommand command;

            // Create the SelectCommand.
            command = new OleDbCommand( "SELECT * FROM Customers " +
                "WHERE Country = ? AND City = ?", connection );

            command.Parameters.Add( "Country", OleDbType.VarChar, 15 );
            command.Parameters.Add( "City", OleDbType.VarChar, 15 );

            adapter.SelectCommand = command;

            // Create the InsertCommand.
            command = new OleDbCommand(
                "INSERT INTO Customers (CustomerID, CompanyName) " +
                "VALUES (?, ?)", connection );

            command.Parameters.Add(
                "CustomerID", OleDbType.Char, 5, "CustomerID" );
            command.Parameters.Add(
                "CompanyName", OleDbType.VarChar, 40, "CompanyName" );

            adapter.InsertCommand = command;
            return adapter;
        }

        private static OleDbDataAdapter CreateCustomerAdapter(
            OleDbConnection connection )
        {
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
            OleDbCommand command;
            OleDbParameter parameter;

            // Create the SelectCommand.
            command = new OleDbCommand( "SELECT * FROM dbo.Customers " +
                "WHERE Country = ? AND City = ?", connection );

            command.Parameters.Add( "Country", OleDbType.VarChar, 15 );
            command.Parameters.Add( "City", OleDbType.VarChar, 15 );

            dataAdapter.SelectCommand = command;

            // Create the UpdateCommand.
            command = new OleDbCommand(
                "UPDATE dbo.Customers SET CustomerID = ?, CompanyName = ? " +
                "WHERE CustomerID = ?", connection );

            command.Parameters.Add(
                "CustomerID", OleDbType.Char, 5, "CustomerID" );
            command.Parameters.Add(
                "CompanyName", OleDbType.VarChar, 40, "CompanyName" );

            parameter = command.Parameters.Add(
                "oldCustomerID", OleDbType.Char, 5, "CustomerID" );
            parameter.SourceVersion = DataRowVersion.Original;

            dataAdapter.UpdateCommand = command;

            return dataAdapter;
        }

        public void SourcesToAdapter()
        {
            // Assumes that customerConnection is a valid SqlConnection object.  
            // Assumes that orderConnection is a valid OleDbConnection object.  
            SqlDataAdapter custAdapter = new SqlDataAdapter(
                "SELECT * FROM dbo.Customers", customerConnection);
            OleDbDataAdapter ordAdapter = new OleDbDataAdapter(
                "SELECT * FROM Orders", orderConnection);

            DataSet customerOrders = new DataSet();

            custAdapter.Fill( customerOrders, "Customers" );
            ordAdapter.Fill( customerOrders, "Orders" );

            DataRelation relation = customerOrders.Relations.Add("CustOrders",
                customerOrders.Tables["Customers"].Columns["CustomerID"],
                customerOrders.Tables["Orders"].Columns["CustomerID"]);

            foreach ( DataRow pRow in customerOrders.Tables[ "Customers" ].Rows )
            {
                Console.WriteLine( pRow[ "CustomerID" ] );
                foreach ( DataRow cRow in pRow.GetChildRows( relation ) )
                    Console.WriteLine( "\t" + cRow[ "OrderID" ] );
            }

        }

        public static SqlDataAdapter CreateSqlDataAdapter( SqlConnection connection )
        {
            SqlDataAdapter adapter = new()
            {
                MissingSchemaAction = MissingSchemaAction.AddWithKey,

                    // Create the commands.
                    SelectCommand = new SqlCommand(
                        "SELECT CustomerID, CompanyName FROM CUSTOMERS", connection),
                    InsertCommand = new SqlCommand(
                        "INSERT INTO Customers (CustomerID, CompanyName) " +
                        "VALUES (@CustomerID, @CompanyName)", connection),
                    UpdateCommand = new SqlCommand(
                        "UPDATE Customers SET CustomerID = @CustomerID, CompanyName = @CompanyName " +
                        "WHERE CustomerID = @oldCustomerID", connection),
                    DeleteCommand = new SqlCommand(
                        "DELETE FROM Customers WHERE CustomerID = @CustomerID", connection)

                };

            // Create the parameters.
            adapter.InsertCommand.Parameters.Add( "@CustomerID",
                SqlDbType.Char, 5, "CustomerID" );
            adapter.InsertCommand.Parameters.Add( "@CompanyName",
                SqlDbType.VarChar, 40, "CompanyName" );

            adapter.UpdateCommand.Parameters.Add( "@CustomerID",
                SqlDbType.Char, 5, "CustomerID" );
            adapter.UpdateCommand.Parameters.Add( "@CompanyName",
                SqlDbType.VarChar, 40, "CompanyName" );
            adapter.UpdateCommand.Parameters.Add( "@oldCustomerID",
                SqlDbType.Char, 5, "CustomerID" ).SourceVersion =
                DataRowVersion.Original;

            adapter.DeleteCommand.Parameters.Add( "@CustomerID",
                SqlDbType.Char, 5, "CustomerID" ).SourceVersion =
                DataRowVersion.Original;

            return adapter;
        }

        public void TheSameForOdbcAndOledb()
        {
            string selectSQL =
                "SELECT CustomerID, CompanyName FROM Customers " +
                "WHERE CountryRegion = ? AND City = ?";
            string insertSQL =
                "INSERT INTO Customers (CustomerID, CompanyName) " +
                "VALUES (?, ?)";
            string updateSQL =
                "UPDATE Customers SET CustomerID = ?, CompanyName = ? " +
                "WHERE CustomerID = ? ";
            string deleteSQL = "DELETE FROM Customers WHERE CustomerID = ?";
        }

        public void OledbExample()
        {
            string selectSQL =
                "SELECT CustomerID, CompanyName FROM Customers " +
                "WHERE CountryRegion = ? AND City = ?";
            string insertSQL =
                "INSERT INTO Customers (CustomerID, CompanyName) " +
                "VALUES (?, ?)";
            string updateSQL =
                "UPDATE Customers SET CustomerID = ?, CompanyName = ? " +
                "WHERE CustomerID = ? ";
            string deleteSQL = "DELETE FROM Customers WHERE CustomerID = ?";
            // Assumes that connection is a valid OleDbConnection object.  
            OleDbDataAdapter adapter = new OleDbDataAdapter();

            OleDbCommand selectCMD = new OleDbCommand(selectSQL, connection);
            adapter.SelectCommand = selectCMD;

            // Add parameters and set values.  
            selectCMD.Parameters.Add(
              "@CountryRegion", OleDbType.VarChar, 15 ).Value = "UK";
            selectCMD.Parameters.Add(
              "@City", OleDbType.VarChar, 15 ).Value = "London";

            DataSet customers = new DataSet();
            adapter.Fill( customers, "Customers" );


            // odbc example with these sqlStrings
            // Assumes that connection is a valid OdbcConnection object.  
            OdbcDataAdapter adapter = new OdbcDataAdapter();

            OdbcCommand selectCMD = new OdbcCommand(selectSQL, connection);
            adapter.SelectCommand = selectCMD;

            //Add Parameters and set values.  
            selectCMD.Parameters.Add( "@CountryRegion", OdbcType.VarChar, 15 ).Value = "UK";
            selectCMD.Parameters.Add( "@City", OdbcType.VarChar, 15 ).Value = "London";

            DataSet customers = new DataSet();
            adapter.Fill( customers, "Customers" );
        }


        public void FillWithSchema()
        {
            // first way
            var custDataSet = new DataSet();

            custAdapter.FillSchema( custDataSet, SchemaType.Source, "Customers" );
            custAdapter.Fill( custDataSet, "Customers" );
            // second way
            var custDataSet = new DataSet();

            custAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
            custAdapter.Fill( custDataSet, "Customers" );

        }

        public void OrderInsetrUpdateDelete()
        {
            DataTable table = dataSet.Tables["Customers"];

            // First process deletes.
            adapter.Update( table.Select( null, null, DataViewRowState.Deleted ) );

            // Next process updates.
            adapter.Update( table.Select( null, null,
              DataViewRowState.ModifiedCurrent ) );

            // Finally, process inserts.
            adapter.Update( table.Select( null, null, DataViewRowState.Added ) );
        }

        public static void BatchUpdate( DataTable dataTable, Int32 batchSize )
        {
            // Assumes GetConnectionString() returns a valid connection string.
            string connectionString = GetConnectionString();

            // Connect to the AdventureWorks database.
            using ( SqlConnection connection = new
              SqlConnection( connectionString ) )
            {

                // Create a SqlDataAdapter.
                SqlDataAdapter adapter = new SqlDataAdapter();

                // Set the UPDATE command and parameters.
                adapter.UpdateCommand = new SqlCommand(
                    "UPDATE Production.ProductCategory SET "
                    + "Name=@Name WHERE ProductCategoryID=@ProdCatID;",
                    connection );
                adapter.UpdateCommand.Parameters.Add( "@Name",
                   SqlDbType.NVarChar, 50, "Name" );
                adapter.UpdateCommand.Parameters.Add( "@ProdCatID",
                   SqlDbType.Int, 4, "ProductCategoryID" );
                adapter.UpdateCommand.UpdatedRowSource = UpdateRowSource.None;

                // Set the INSERT command and parameter.
                adapter.InsertCommand = new SqlCommand(
                    "INSERT INTO Production.ProductCategory (Name) VALUES (@Name);",
                    connection );
                adapter.InsertCommand.Parameters.Add( "@Name",
                  SqlDbType.NVarChar, 50, "Name" );
                adapter.InsertCommand.UpdatedRowSource = UpdateRowSource.None;

                // Set the DELETE command and parameter.
                adapter.DeleteCommand = new SqlCommand(
                    "DELETE FROM Production.ProductCategory "
                    + "WHERE ProductCategoryID=@ProdCatID;", connection );
                adapter.DeleteCommand.Parameters.Add( "@ProdCatID",
                  SqlDbType.Int, 4, "ProductCategoryID" );
                adapter.DeleteCommand.UpdatedRowSource = UpdateRowSource.None;

                // Set the batch size.
                adapter.UpdateBatchSize = batchSize;

                // Execute the update.
                adapter.Update( dataTable );
            }
        }



    }

}

/*
ASP.NET application:

Web.config Configuration:
To enable impersonation for all requests in an ASP.NET application, you can modify the web.config file. Add the <identity> element under the <system.web> section and set the impersonate attribute to true.
<system.web>
    <identity impersonate="true"/>
</system.web>

To impersonate a specific user for all requests, you can specify the userName and password attributes in the <identity> element.
<system.web>
    <identity impersonate="true" userName="username" password="password"/>
</system.web>

Column types are created as .NET Framework types according to the tables in Data Type Mappings 
in ADO.NET. Primary keys are not created unless they exist in the data source and DataAdapter.
MissingSchemaAction is set to MissingSchemaAction.AddWithKey. 
If Fill finds that a primary key exists for a table, it will overwrite data in the DataSet 
with data from the data source for rows where the primary key column values match those 
of the row returned from the data source.

// Assumes that connection is a valid SqlConnection object.  
string queryString =
  "SELECT CustomerID, CompanyName FROM dbo.Customers";  
SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);  
  
DataSet customers = new DataSet();  
adapter.Fill(customers, "Customers");


 */