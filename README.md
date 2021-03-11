# UtilityLibrary

## Table of Contents
* [Credential API](Credential-API)
* [DataBaseHandler](DatabaseHandler)
* [DateTime](DateTime)
* [Excel](Excel)
* [Loggers](Loggers)
* [ProcessManager](ProcessManager)

## Summary
This is a library which will include varies classes that the user will be able to import and use in their projects.

## Credential-API
The credential API is a class that connects to a Rest API Server. This server lets you'll retrieve usernames and passwords that are stored in a database. This is useful since you wont have to store the passwords in the projects you have and will only have to update the password in the [Credential Repo application](https://github.azc.ext.hp.com/SLAC-Dev/credential-app).

## DataBaseHandler

This class DataBaseHandler allows the user to setup a connection to a database and pass sql queries. To use the class correctly you'll have to follow the following steps:
* 1- Instantiate the `DataBaseHandler` Class in a object.
* 2- Assign the `_ConnectionString` with the database conexion that you wish to conect to.
* 3- Use the `OpenConnection()` to setup a connexion. 
* 4- Use the `SetSQLQuery()` which recieves a SQL command. So basically pass any type of command that you would use in SQL Server, even Stored Procedure commands.
* 5- After this there are 2 options you can use either the `executeQuery` function or the `execute` function. The `executeQuery` is used when you wish to retrieve that info so queries which use the `Select` command are the only ones that will work or any other commands that display table info. The info will be saved to a `List<List<object>>`. The other function that you can use is called `execute` this is perfect when you want to execute commands that don't return info, for example update, delete, or sql stored procedures.

## DateTime
This class has 2 functions one that allows you to get the central standard time and one that converts a date that you pass by parameter to central standard time.
* `getCentralStandardTime()`
* `convertToCentralStandardTime(DateTime date)`

## Excel
This Folder includes 3 class:
* CSVReader: This class enables you to read csv files and store the info in a DataTable object.
* ExcelReader: This class allows you to read excel files and create dictionary from `worksheets`, `object[,]`, `List<string[]>` or `string[]`
* OledbReader: This is another class that allows you to read excel files using the `OleDbConnection`.

## Loggers
This Folder includes 4 classes:
* AppLogger: This class allows you to log info to System.Windows.Forms applications.
* ConsoleLogger: This class allows you to log info to the console. Errors are logged in the color red, importante messages are logged in yellow and regular messages are logged in white.
* Logger: Error messages are stored in file that the user can create with the Logger class.
* TextLogger: This class acts like a stracer so you can log whatever you want in a file that the user can create with the Logger class.

## ProcessManager 
This class allows you to kill processes that are running like excel files that don't close automatically.



