## <u>1. Overview: DbaseFrame</u>
Project to gain usage of Excel-,Access-  and Sql-data with C#.

My 'NPOIwrap' is closed for the reason that i believe 'NPOI' is <u>unsafe</u>. While debugging one of my programs these add-ons from 'NPOI' couldn't be deleted - they were hooked into the system and where denying their removal that way. You have to wake up to these hacks and move into action. For me it means starting  a new project with good old topics. The book 'C# 12 in a nutshell' claims to be a good source for this ambition ... so here i am today 15. October 2024.

Starting as example codes for my programs i plan to create classes that handle it all. I know understanding something is meaning you are not needing it - but you use OOPs for that convenience.

## <u>2. Excel file handling</u>

I will use **'System.Data.OleDb'** that is installed as NuGet package.

They give you an easy job, if you are able of SQL. There are the drawbacks in this procedure and i personally don't like to have to form a string in this SQL way. First argument to finish this once for all into a helper class.

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
            temp[ pos ] = reader[ pos ].ToString();
        values.Add( temp );

    }
            
}
```

In this example the whole Excel spreadsheet is read in as a list of string[] - one array position for every column. Purpose is the buffering of anonymous data. You wouldn't need this on a known data constellation and could use Entity Framework for convenience.
