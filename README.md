available on nuget.org as package '**DbaseFrame**' ( this README describes package version **1.0.8**, state 27.11.2026, **Update 28.11.2026: 1.0.9**; see chapter 5 for the current behaviour ).



## <u>1. Overview: DbaseFrame</u>

Project to gain usage of Excel-,Access-  and Sql-data with C#. This project serves as **example WPF-program** and **source** for the **utility classes**.

**Update 27.11.2026:** SQL is not part of the package: there was no code for it, so the unused packages 'Microsoft.Data.SqlClient' and 'System.Data.Odbc' and the unfinished class 'ExcelWriteStringList' were removed. The package now needs only 'System.Data.OleDb' ( 6 instead of 19 DLLs ).

My 'NPOIwrap' is closed for the reason that i believe 'NPOI' is <u>unsafe</u>. While debugging one of my programs these add-ons from 'NPOI' couldn't be deleted - they were hooked into the system and where denying their removal that way. You have to wake up to these hacks and move into action. For me it means starting  a new project with good old topics. The book 'C# 12 in a nutshell' claims to be a good source for this ambition ... so here i am today 15. October 2024.

Starting as example codes for my programs i plan to create classes that handle it all. I know understanding something is meaning you are not needing it - but you use OOPs for that convenience.

**Standard disclaimer ( updated 27.11.2026 ): writing with OleDb works for <u>simple</u> tables only – one table in a <u>new</u> file, numbers and texts up to 255 characters. For anything else ( updating existing tables, formats, several tables ) use <u>Interop</u>.**

## <u>2. Excel file handling with OleDb works, but it's not good enough</u>

**Standard disclaimer ( updated 27.11.2026 ): writing with OleDb works for <u>simple</u> tables only – one table in a <u>new</u> file, numbers and texts up to 255 characters. For anything else ( updating existing tables, formats, several tables ) use <u>Interop</u>.**

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

**Update 27.11.2026:** the cause is found and fixed. Closing and reopening the connection between 'CREATE TABLE' and 'INSERT INTO' makes the ACE driver store every number as text. Keeping the connection open writes real numbers ( checked in the xlsx file itself ).

Second problem for correctness is the produced table from writing is seen in Excel as one table ( good ) but if you read them in with this code two tables are found. The next barrier for the usability of **OleDb**.

**Update 27.11.2026:** the reason: 'CREATE TABLE' makes the ACE driver create a sheet 'newStrings' **and** a named range 'newStrings' on it ( seen in the xlsx file ). Reading lists both: 'newStrings' ( the named range ) and 'newStrings$' ( the sheet ). Both deliver the same data, you can choose either.

I will let this code stay as good example but only for the interest. You could use it as dirty version of data management but **Interop** there and many people have the Office suite via their OneDrive like i do. 

## <u>2.1 demoprogram's menu 'OleDb Excel'</u>

An example procedure for testing the class is given here.

1. '**open Excel file by dialog ( no headers )/( with headers )'** lets you choose the right file with the common file dialog. If you want to change it you choose again. Any operation afterwards uses that chosen Excel file.
2. **'read the tables'** gives you the dialog to choose one of the found table in this file.
3. **'read table names by number'** show how to query the table's name with a index number.
4. **'read chosen table as List of string[]'** will read any cell of the table as a 'string'.
5. **'read the chosen table as List of double[]'** will read the cells as double if they are of that type or you will see a 0 ( standard initialization ). Good for buffering data in with no exception.
   **Update 27.11.2026:** the data have to fit now: empty cells give 0, TRUE/FALSE 1/0, a number as text ( "1,5" ) is converted, but any other text gives an error message with row and column. Before, text was silently 0.
6. **'write the read double-list into a new file'** will create a new file with the data in it.
7. **'write the read string-list into a new file'** will do the same for the read string values.
   **Update 27.11.2026:** both report the result ( written / nothing written ), a cancelled dialog writes nothing, and with headers the read column names become the column names of the new table.

The demo is using one instance of the **'ExcelOleDb'** class for the whole show. While i use internally lists for the read data you easily can use arrays as they are sort of ambiguous towards each other.

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

**Standard disclaimer ( updated 27.11.2026 ): writing with OleDb works for <u>simple</u> tables only – one table in a <u>new</u> file, numbers and texts up to 255 characters. For anything else ( updating existing tables, formats, several tables ) use <u>Interop</u>.**

Here you can see the full ability of the device driver givers. As good as nothing had to be changed in the script except the connection string which points to the Access file.

That means the content of the second chapter counts here, too. Leaving the question if the Access data base will take doubles as double or not. But this topic is useless as i could not find a way to read the values in like they are there in the Access file. Writing would be the second step in this setup - but where is my mistake?.

Having checked the connection strings is easy game and tells you that the way of fetching the data ( with a reader or from a data adapter ) doesn't count much. Important would be the data **coming in** and there i can show no results. The data is just not there except the shape of the spreadsheets - hmm.

I found out that my Access version ( from the One-Drive-Abo ) is not taking floating numbers. Even worse my numbers imported from Excel are made to 'string'. This explains why i found only the shape but not the data. **As slow learner** i sure tried to import again and it shows that any fast action is giving me this mistake. Errors that any beginner does: being too fast and trusting where corrected and now the software is reading the data set correctly with '**OleDb**'. That means i have to tell exactly which column has which data type on import and then it works.

That showed up one limitation of Access: you can't change a field type after creating the table. A wrong import locks your whole data set .

I read that Microsoft did drop OleDb years ago and didn't continue to develop it further. But with .NET 8.x it is still there and thus should work ( **update:** the project now targets .NET 10, 'System.Data.OleDb' 11.0.0, still working ). While the shortcut ACE in the drivers name would mean 'Access Connectivity Engine' it is long from working - why do they add it to the system then?

For completeness i let the stuff stay here.

**Update 27.11.2026:** reading Access works ( 'AccessOleDb', tested with the files 'Access_Test.accdb' and '.mdb' ): every number type ( also Int16, Decimal, Byte, date ) is read, text in a number column gives an error message, the table choice shows only your own tables ( no 'MSys...' system tables ). Writing works too: the ACE driver cannot create a new Access file via OleDb ( "file not found" ), so the new file is created with **ADOX** ( comes with the ACE driver ) and then filled. Access keeps trailing spaces in texts, Excel via ACE does not.



### <u>4. Excel Interop</u>

With my Visual Studio Pro i am now allowed to use the interop functionality. As i have the One Drive Personal the usage of Office is part of my system. Something to keep in mind wondering about the why i want to use interop exclusively in the future.

Here you can use native c# scripts to reach Office and it does look good. But beware an added 'using ...' can give you an enormous amount of errors. Trying to use the com-classes for office shows that Office must be scripted with a visual language from Microsoft. It starts with the UI's partial class 'MainWindow' derived from 'Window'. Both the Excel and the Word com classes have a 'Window' too, leading to an amount of errors. This problem made me drop the 'using ... '-statement for the interop classes. 

I even think you can't change that decision for any other projects because of the debugger being lazy. Better stay secure and working than ignoring possible mistakes.

```c#
using Excel = Microsoft.Office.Interop.Excel;
```

You need to add the com-type to the dependencies:

​	"Microsoft Excel 16.0 Object Library"

later you can have your fields in the program:

```c#
        Microsoft.Office.Interop.Excel.Application excelApp = 
            new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook workbook;
		Microsoft.Office.Interop.Excel.Worksheet worksheet;
```

### <u>4.1 my inplementation</u>

... is prepared differently than the former parts of this project. 

Because i can use this without the ugly SQL-strings i just went to a different organisation of the class.

The bigger part works on a chosen or given workbook - the second smaller part just stores data in a new one.

The try/catch-block is needed at the beginning when opening the file ( **update 27.11.2026:** the reason of a failed opening is kept in 'lastError': missing, locked, damaged, password protected; and every possible question of the invisible Excel is switched off beforehand, otherwise the program would wait forever ):

```c#
            ExcelInterop excel = new ExcelInterop( "", false, doUseHeader: true, doReadOnly: true );
            if ( !excel.workbookOpen )
                Console.WriteLine( "not opened: " + excel.lastError );
```

Like before you can read the contents of the excel file as type-, string- or double into the classes internal fields but this time i use 2d arrays. For the safety every cell needs to be checked if the type is the correct one. **Update 27.11.2026:** all values of the data rectangle are fetched with **one** access ( 2000 x 5 cells: 110 ms instead of 23 s cell by cell ) and checked afterwards one by one. The arrays are **Matrix( x, y ) = point( x, y )**: first index column, second index row; Excel's own order [ row, column ] is only used at the hand-over.

```c#
            excel.SelectSheet( 0 );          // sheets count from 0
            excel.ReadDoubles( );            // text instead of a number -> error message with cell address
            double[,] data = excel.valuesDoubleArray;   // data[ x, y ]
            string[] names = excel.headers;              // the header row ( with doUseHeader )
```

In writing the data to an excel workbook the strength of the interop classes can be seen: one assignment of 'Value2' for the whole range. **Update 27.11.2026:** 'WriteDoublesToNewTarget' / 'WriteStringsToNewTarget' really create and save the new file ( with the header row, if read with headers ), return false if the dialog is cancelled, and Excel is ended completely by 'ExcelInteropQuit( )' ( every COM object is released, otherwise Excel keeps running ).



### <u>5. State 27.11.2026 ( package 1.0.8 )</u>

**The data have to fit, otherwise there is an error message** – nothing is bent silently:

- an empty sheet, a sheet with only a header row, an empty first/last row or column of the data rectangle
- text where a number is expected ( with cell address or row and column )
- NaN or infinity when writing, texts longer than 255 characters when writing with OleDb
- empty cells inside the data become "" or 0, TRUE/FALSE become 1/0, Excel errors ( #DIV/0! ) become 0 or their text

**Reading is all or nothing:** if a read fails, its result field is empty ( not half filled, not the data of an earlier read ), and writing without data reports it.

**Files:** missing, damaged, password protected or locked files give an understandable message. A file locked completely by another program used to crash the whole program inside the ACE driver; this is now checked beforehand.

**Classes and versions:** 'ExcelInterop' 1.0.8 ( Excel via COM, needs Excel ), 'ExcelOleDb' 1.0.14 and 'AccessOleDb' 1.0.6 ( ACE driver ). The package number is kept in the project file; 'PackageInfo' delivers version and commit, the demo program shows them at the start.

**Known limit of the ACE driver:** it guesses the type of a column from the first 8 rows. Text appearing later in a number column comes in as empty ( 0 ). For strict checking of every cell use 'ExcelInterop'.

**Demo program:** menus 'OleDb Excel', 'Interop Excel' ( 14 items ) and 'OleDb Access'; every error is shown in the window and the program goes on. Test files: 'Parable_Demo.xlsx', 'StringGolfer.xlsx' ( texts ), 'Defekt_Test.xlsx' ( damaged ), 'Access_Test.accdb'.

**Update 28.11.2026 ( package 1.0.9 ):** all three classes offer 'ReadColumnNames( int numTable )': the column names of a table or sheet without reading the data, the same names a read puts into 'headers'. A number that does not exist and an empty sheet give an error message. 'AccessOleDb.ColumnNameDummy( )' did nothing and is marked obsolete. 'ExcelOleDb' now counts only worksheets ( no named ranges like the print area or an autofilter ); its numbers follow the alphabetical order of the ACE driver, while 'ExcelInterop' counts in the order of the sheets in Excel. Versions: 'ExcelInterop' 1.0.9, 'ExcelOleDb' 1.0.15, 'AccessOleDb' 1.0.7. New test files: 'Blaetter_Test.xlsx' ( several sheets, empty sheet, header only, duplicate headers, mixed column ), 'Leer_Test.xlsx' and damaged files ( 'Null_Test.xlsx', 'KeinZip_Test.xlsx', 'ZipOhneMappe_Test.xlsx', 'XmlKaputt_Test.xlsx' ).

**Update 28.11.2026 ( limits of the ACE driver behind 'ExcelOleDb', measured ):** the driver guesses the type of a column from its first 8 rows. Text that comes later in a number column ( 'no number', TRUE, '1,5' ) arrives in 'ReadStringList' as an empty text, without an error. A damaged sheet inside the workbook can be read up to the damaged spot without an error ( 'XmlKaputt_Test.xlsx': 3 of 10 rows ). 'ExcelInterop' takes every cell value from Excel itself and does not guess column types, so it suits mixed columns better.

**Update 28.11.2026 ( writing without message boxes from the library ):** the classes no longer show a message box of their own, so they can also be used where nobody can click ( service, server ). All six 'Write…ToNewTarget' functions have a new last parameter 'overwrite' ( default false ). Writing without data throws an 'InvalidOperationException'. An existing target file is replaced only if it was confirmed in the save dialog ( Windows asks there itself ) or 'overwrite: true' was passed; otherwise an 'IOException' is thrown and the file stays untouched. A target file that cannot be replaced ( opened in another program, read-only ) also gives an 'IOException'. The new file is first written to a temporary file in the same folder and replaces the old one only when writing has succeeded: if anything fails while writing, the old file is unchanged ( before, it was deleted first and could be lost ). 'ExcelInterop' refuses the workbook that is open as target ( Excel cannot save a new workbook under its name ). The return value false now means only: the dialog was cancelled. The file dialogs themselves are still part of the classes. A folder where no file may be created ( missing, no write permission ) gives an 'IOException' that names the folder.

**Update 28.11.2026 ( table names when writing with OleDb, measured with the ACE driver ):** the name of the new table is checked before the dialog and an invalid one gives an 'ArgumentException' with the reason ( before, the driver only said 'syntax error in CREATE TABLE' ). Allowed are at most 64 characters, no leading space, and not the characters ! ` [ ] plus : $ for Excel and . for Access. Note for Excel: the driver creates the sheet with a changed name – spaces and other special characters become _, a dot becomes #, a leading digit gets a _ in front, more than 31 characters are cut off ( 'a b' becomes sheet 'a_b' ). Access keeps the name as given.

### <u>14.Donations</u>

You can if you want donate to me for the **GitHub content**. I always can use it, thank you.

https://www.paypal.com/ncp/payment/QBF7E2ZG4J8NU



