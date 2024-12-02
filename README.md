available as '**DbaseFrame.1.0.3.nupkg**'.



## <u>1. Overview: DbaseFrame</u>

Project to gain usage of Excel-,Access-  and Sql-data with C#. This project serves as **example WPF-program** and **source** for the **utility classes**.

My 'NPOIwrap' is closed for the reason that i believe 'NPOI' is <u>unsafe</u>. While debugging one of my programs these add-ons from 'NPOI' couldn't be deleted - they were hooked into the system and where denying their removal that way. You have to wake up to these hacks and move into action. For me it means starting  a new project with good old topics. The book 'C# 12 in a nutshell' claims to be a good source for this ambition ... so here i am today 15. October 2024.

Starting as example codes for my programs i plan to create classes that handle it all. I know understanding something is meaning you are not needing it - but you use OOPs for that convenience.

## <u>2. Excel file handling with OleDb works, but it's not good enough</u>

I will use **'System.Data.OleDb'** that is installed as NuGet package.

Where do you need this and not **Entity Framework** ? Anonymous array of data wouldn't be so easy. An Excel spreadsheet can be taken as a list of rows that contain the array data. If you use **Entity Framework** you need to know the data if you want to read them or you use a list and **EF** would create two tables for the basic data class and then for the members of the list data. So a list can be managed by **EF** but it would be in connection with a second table having every list member in an entry there. 

Buffering in an array of 'array data' like a spreadsheet looks can be done direct. That data handling will always be different than you think - SQL is used everywhere in **OleDb** and **EF**.  SQL is demanding a clear definition of the used columns if you create a table or add data to it.

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

**<u>Why or what is not working?</u>**

Enough documentation is not to be found. AI answers change from day to day and you can have a good time testing that out.

Very annoying is the fact that written 'doubles' are not recognized as 'doubles' from Excel but they seem to be 'string'. I couldn't find the solution for that problem.

Second problem for correctness is the produced table from writing is seen in Excel as one table ( good ) but if you read them in with this code two tables are found. The next barrier for the usability of **OleDb**.

I will let this code stay as good example but only for the interest. You could use it as dirty version of data management but i hope for the **OpenXML-SDK**.

## <u>2.1 demoprogram's menu 'OleDb Excel'</u>

An example procedure for testing the class is given here.

1. '**open Excel file by dialog ( no headers )/( with headers )'** lets you choose the right file with the common file dialog. If you want to change it you choose again. Any operation afterwards uses that chosen Excel file.
2. **'read the tables'** gives you the dialog to choose one of the found table in this file.
3. **'read table names by number'** show how to query the table's name with a index number.
4. **'read chosen table as List of string[]'** will read any cell of the table as a 'string'.
5. **'read the chosen table as List of double[]'** will read the cells as double if they are of that type or you will see a 0 ( standard initialization ). Good for buffering data in with no exception.
6. **'write the read double-list into a new file'** will create a new file with the data in it.
7. **'write the red strin-list into a new file'** will do the same for the read string values.

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

The example is presented with an Excel source having no headers while the writing routines use them. Remember to read written files with headers or you will loose one row of data.

### <u>2.2 Excel and the header rows</u>

It is said that there is a line 0 being used by Excel for the column names - but it is hidden generally. And there is no real usage for them to be set in Excel by hand. Most people just use line 1 for the header row. This way the driver sees it, too.

Working magic on this OleDb-driver and Entitiy Framework is SQL. <u>But SQL is not able of magic</u>. Thus there is no real intelligent way to find a header row. You just have to query the user for it or set **'HDR = YES;'** or **'HDR = NO;'** in the connection string like you want. Maximum consequence would be the loss of one row of data in reading.

Writing is a different situation, as SQL demands a name for every column in your becoming Excel spreadsheet. So you generally write files with **'HDR = YES;'** option in your connection string.

I personally have enough with pushing an array into an Excel spreadsheet - one table in one file. Always a clean file for the data. If you need more versatility in writing and updating tables in a file ask for it and maybe i can come up with something.

### <u>3. Access with OleDb is even worse</u>

Here you can see the full ability of the device driver givers. As good as nothing had to be changed in the script except the connection string which points to the Access file.

That means the content of the second chapter counts here, too. Leaving the question if the Access data base will take doubles as double or not. But this topic is useless as i could not find a way to read the values in like they are there in the Access file. Writing would be the second step in this setup - but where is my mistake?.

Having checked the connection strings is easy game and tells you that the way of fetching the data ( with a reader or from a data adapter ) doesn't count much. Important would be the data **coming in** and there i can show no results. The data is just not there except the shape of the spreadsheets - hmm.

I found out that my Access version ( from the One-Drive-Abo ) is not taking floating numbers. Even worse my numbers imported from Excel are made to 'string'. This explains why i found only the shape but not the data. **As slow learner** i sure tried to import again and it shows that any fast action is giving me this mistake. Errors that any beginner does: being too fast and trusting where corrected and now the software is reading the data set correctly with '**OleDb**'. That means i have to tell exactly which column has which data type on import and then it works.

That showed up one limitation of Access: you can't change a field type after creating the table. A wrong import locks your whole data set .

I read that Microsoft did drop OleDb years ago and didn't continue to develop it further. But with .NET 8.x it is still there and thus should work. While the shortcut ACE in the drivers name would mean 'Access Connectivity Engine' it is long from working - why do they add it to the system then?

For completeness i let the stuff stay here.



### <u>14.Donations</u>

You can if you want donate to me for the **GitHub content**. I always can use it, thank you.

https://www.paypal.com/ncp/payment/QBF7E2ZG4J8NU



