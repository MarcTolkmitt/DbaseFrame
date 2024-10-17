## <u>1. Overview: DbaseFrame</u>
Project to gain usage of Excel-,Access-  and Sql-data with C#.

My 'NPOIwrap' is closed for the reason that i believe 'NPOI' is <u>unsafe</u>. While debugging one of my programs these add-ons from 'NPOI' couldn't be deleted - they were hooked into the system and where denying their removal that way. You have to wake up to these hacks and move into action. For me it means starting  a new project with good old topics. The book 'C# 12 in a nutshell' claims to be a good source for this ambition ... so here i am today 15. October 2024.

Starting as example codes for my programs i plan to create classes that handle it all. I know understanding something is meaning you are not needing it - but you use OOPs for that convenience.

## <u>2. Excel file handling</u>

I will use **'System.Data.OleDb'** that is installed as NuGet package.

Where do you need this and not Entity Framework ? Anonymous array of data wouldn't be so easy. An Excel spreadsheet can be taken as a list of rows that contain the array data. If you use Entity Framework you need to know the data if you want to read them or you use a list and EF would create two tables for the basic data class and then for the members of the list data. So a list can be managed by EF but it would be in connection with a second table having every list member in an entry there. 

Buffering in an array of 'array data' like a spreadsheet looks can be done direct. That data handling will always be different than you think - SQL is used everywhere in OleDb and EF.  SQL is demanding a clear definition of the used columns if you create a table or add data to it.

They give you an easy job, if you are able of SQL - the drawbacks in this procedure. And i personally don't like to have to form a string in this SQL way. First argument to finish this once and for all into a helper class.

Procedure is always the same in having a 'connectionString' like 

```c#
connectionString = 
    "Provider=Microsoft.ACE.OLEDB.12.0;" +
    "Data Source=" +
    "C:\\" + 
    "Parable_Demo.xlsx" +
	";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";
```

Here you have the provider, the source file and special parameters in one set.

And then you open a connection and send your SQL command like

```c#
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
            temp[ pos ] = reader.GetString( pos );
        values.Add( temp );

    }
            
}
```

In this example the whole Excel spreadsheet is read in as a list of string[] - one array position for every column. Purpose is the buffering of anonymous data. You wouldn't need this on a known data constellation and could use Entity Framework for convenience.
