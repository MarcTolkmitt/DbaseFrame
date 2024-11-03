available as '**DbaseFrame.1.0.0.nupkg**'.



## <u>1. Overview: DbaseFrame</u>

Project to gain usage of Excel-,Access-  and Sql-data with C#. This project serves as **example WPF-program** and **source** for the **utility classes**.

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
            temp[ pos ] = reader[ pos ].ToString();
        values.Add( temp );

    }
            
}
```

In this example the whole Excel spreadsheet is read in as a list of string[] - one array position for every column. Purpose is the buffering of anonymous data. You wouldn't need this on a known data constellation and could use Entity Framework for convenience.

## <u>2.1 demoprogram's menu 'OleDb Excel'</u>

An example procedure for testing the class is given here.

1. '**open Excel file by dialog'** lets you choose the right file with the common file dialog. If you want to change it you choose again. Any operation afterwards uses that chosen Excel file.
2. **'read the tables'** gives you the dialog to choose one of the found table in this file.
3. **'read table names by number'** show how to query the table's name with a index number.
4. **'read chosen table as List of string[]'** will read any cell of the table as a 'string'.
5. **'read the chosen table as List of double[]'** will read the cells as double if they are of that type or you will see a 0 ( standard initialization ). Good for buffering data in with no exception.

The demo is using one instance of the **'DbaseFrameExcel'** class for the whole show. While i use internally lists for the read data you easily can use arrays as they are sort of ambiguous towards each other.

```c#
readExcel.ReadStringList();
foreach ( string[] line in readExcel.valuesString )
    Display( ArrayToString( line ) );

var rowArray = readExcel.valuesString.ToArray();
	Display( ArrayJaggedToString( rowArray, true ) );

var listRows = rowArray.ToList();
    foreach ( string[] line in listRows )
        Display( ArrayToString( line ) );
```

=> 'listRows' will be a same sized and same looking list like the original.

### <u>4.Donations</u>

You can if you want donate to me for the **GitHub content**. I always can use it, thank you.

https://www.paypal.com/ncp/payment/QBF7E2ZG4J8NU



## <u>5. affiliate links</u>

New and affordable, BTCMiner a service in the cloud: 

#### https://www.btcminer.vip/21663039

