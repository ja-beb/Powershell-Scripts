<#
    Author_____: Sean Bourg <sean.bourg@gmail.com>
    Date_______: 2018-08-01
    Description: Collection of functions for working with an Access database files.

    This requires that the Microsoft Access Database Engine to be installed.

 #>

 function Open-AccessConnection {
    <# 
      .SYNOPSIS
       Open MS SQL Server connection database file.
 
      .DESCRIPTION
       Open MS SQL Server connection database file. 
 
      .PARAMETER Server
      The server that the database resides on.
 
      .PARAMETER Database
      The name of the database to open once connected.
      
      .PARAMETER Username
      Username to use in connection. If ommitted the user's executing credentials are used.
 
      .PARAMETER Password
      User account password
 
      .PARAMETER Timeout
      The number of sections until connection times out.
 
      .OUTPUTS
      System.Data.SqlClient.SqlConnection
 
      .EXAMPLE
      $database = Open-SqlConnection -Server 'SqlServer' -Database 'MyDatabase';
 
        Open database connection.
     #> 
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory = $true)] [string] $File,
        [Parameter(Mandatory = $false)] [SecureString] $Password = $null,
        [Parameter(Mandatory = $false)] [int] $Timeout = 15 
    );
 
    $security = if ([string]::IsNullOrEmpty($Password) ) { 
        'Persist Security Info=False;';
    }
    else {
        'Jet OLEDB:Database Password={0}' -f $Password;        
    } 

    $connectionString = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source={0};{1}" -f $File, $security;
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString);
    $connection.Open();
    return $connection;
}
 
function Get-AccessData {
    <# 
      .SYNOSPIS 
       Get SQL data from database using inputed query.
 
      .DESCRIPTION
       Invoke access database query. Allows for the inclusion of database parameters for binding.
 
      .PARAMETER Connection
       Sql Server database connection to close.
   
      .PARAMETER Query
       SQL Query to execute.
  
      .PARAMETER ParameterList
       A hash containing the SQL parameter in (name,value) pairs.
  
      .OUTPUTS
       System.Data.DataTable containing query results.
 
      .EXAMPLE
       Get-SqlData -Connection $db -Query "SELECT TOP 10 id, username FROM Site.Accounts" | ForEach-Object { Write-Host ( "{0} - {1}" -f $_.id $_.username ); } ;
 
       Query database connection with no parameters.
 
      .EXAMPLE
       Get-SqlData -Connection $db -Query "SELECT id, username FROM Files Where username=@username -Parameters " @{"username" = "my_user";} | ForEach-Object { Write-Host ( "{0} - {1}" -f $_.Id $_.username ); } 
 
       Query database connection with sql parameters.
 
     #>
 
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory = $true)] [System.Data.OleDb.OleDbConnection] $Connection,
        [Parameter(Mandatory = $true)] [string] $Query,
        [Parameter(Mandatory = $false)] [hashtable] $Parameters = @{ }
    );
     
    ## Create query:
    [System.Data.OleDb.OleDbCommand] $command = $Connection.CreateCommand();
    $command.CommandText = $Query;
    $command.CommandTimeout = $null;
         
    ## Load parameters:
    foreach ($key in $Parameters.Keys) {
        [void] $command.Parameters.AddWithValue("@$key", $(if ( $null -eq $Parameters[$key] ) { [DBNull]::Value; } else { $Parameters[$key]; }) );
    }    

    ## Execute query:
    $dataTable = New-Object System.Data.DataTable;
    $dataTable.Load( $command.ExecuteReader() );
    $dataTable;
}
 
function Invoke-AccessQuery {  
    <# 
      .SYNOSPIS 
       Get SQL data from database using inputed query.
 
      .DESCRIPTION
       Invoke access database query. Allows for the inclusion of database parameters for binding.
 
      .PARAMETER Connection
       SQL Server database connection to close.
   
      .PARAMETER Query
       SQL Query to execute.
  
      .PARAMETER ParameterList
       A hash containing the SQL parameter in (name,value) pairs.
  
      .EXAMPLE
       Invoke-SqlQuery -Connection $db -Query "INSERT INTO Accounts(username) VALUES('new_user');";
 
       Execute database query - insert new username.
 
     #>
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory = $true)] [System.Data.OleDb.OleDbConnection] $Connection,
        [Parameter(Mandatory = $true)] [string] $Query,
        [Parameter(Mandatory = $false)] [hashtable] $Parameters = @{ }
    );
     
    [System.Data.OleDb.OleDbCommand] $command = $Connection.CreateCommand();
    $command.CommandText = $Query;
    $command.CommandTimeout = $null;
 
    ## Load parameters:
    foreach ($key in $Parameters.Keys) {
        [void] $command.Parameters.AddWithValue("@$key", $(if ( $null -eq $Parameters[$key] ) { [DBNull]::Value; } else { $Parameters[$key]; }) );
    }
    $command.ExecuteNonQuery();
}
