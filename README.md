# OleDbSql Powershell Module #

## Description ##

A Powershell module that provides a simple means to connect to any OleDB compatible database and execute SQL queries, returning the results as piped objects.

## Features ##

* Tables are returned as System.Management.Automation.PSCustomObject objects, whose type can be specified with the TypeName parameter.
* Scalar results (update, insert, delete, etc) are returned as an integer value representing the quantity of rows affected.
* The returned objects' parameters are automaticly cast from their datbase types to their .Net equivalents.
* Connection strings can be saved with an alias for ease of use.
* Connection strings can contain parameters that query the user for input. Particularly useful for passwords.
* Duplicate column names are automatically given unique property names.
* Uses OleDb.Net. Though OleDb.net is not as fast as a native provider, it is very flexable. For databases with an OleDB.net provider, the same code will work by appending the correct 'Provider' clause to the connection string. See SQLConnections.txt for examples.


## Example ##

```
#!Powershell

    $ConnectionString = 'Provider=sqloledb; Data Source=%Server%; Initial Catalog=%Database%;Integrated Security=SSPI;'
    $ConnectionObject = New-SqlConnection -ConnectionString $ConnectionString
    Invoke-OledbSql -Connection $ConnectionObject -SQL 'select 1 as Ping'

```
This example creates a connection object, querying the user for the database and server names. It then issues a simple select that returns a '1' if the connection succeeds.
