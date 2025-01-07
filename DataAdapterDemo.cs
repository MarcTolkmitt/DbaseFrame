using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
 */