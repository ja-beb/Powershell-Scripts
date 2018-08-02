<#
    Author_____: Sean Bourg <sean.bourg@gmail.com>
    Date_______: 2018-08-01
    Description: Collection of functions for working with an MS SQL Server database.    
 #>

 function Open-SqlConnection
 {
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
             [Parameter(Mandatory=$true)] [string] $Server
                 ,[Parameter(Mandatory=$true)] [string] $Database
                 ,[Parameter(Mandatory=$false)] [string] $Username = $null
                 ,[Parameter(Mandatory=$false)] [SecureString] $Password = $null
                 ,[Parameter(Mandatory=$false)] [int] $Timeout = 15 
         );
 
     BEGIN
     {
         $connectionString = 'Server={0};Database={1};Connect Timeout={2}' -f $Server, $Database, $Timeout;
         if ([string]::IsNullOrEmpty($Username) -eq $true ) 
         { 
             $connectionString = '{0};Integrated Security=True' -f $connectionString;
         }
         else
         {
             $connectionString = '{0};User ID={1};Password={2}' -f $connectionString, $Username, $Password;        
         } 
     }
     
     PROCESS
     { 
         $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString);
         $connection.Open();
         return $connection;
     }
 
     END
     {}
 
 }
 
 function Get-SqlData
 {
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
             [Parameter(Position=0, Mandatory=$true)] [System.Data.SqlClient.SqlConnection] $Connection
                 ,[Parameter(Position=1, Mandatory=$true)] [string] $Query
                 ,[Parameter(Position=2, Mandatory=$false)] [System.Object] $Parameters = @{}
         );
     
     BEGIN
     {
         ## Create query:
         [System.Data.SqlClient.SqlCommand] $command = $Connection.CreateCommand();
         $command.CommandText = $Query;
         $command.CommandTimeout = $null;
         
         ## Load parameters:
         foreach($key in $Parameters.Keys)
         {
             [void] $command.Parameters.AddWithValue("@$key",  $(if ( $null -eq $Parameters[$key] ) { [DBNull]::Value; } else { $Parameters[$key]; }) );
         }    
     }
 
     PROCESS
     {
         ## Execute query:
         $dataTable = New-Object System.Data.DataTable;
         $dataTable.Load( $command.ExecuteReader() );
         $dataTable;
     }
 
     END
     {}
 }
 
 function Invoke-SqlQuery
 { 
 
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
             [Parameter(Position=0, Mandatory=$true)] [System.Data.SqlClient.SqlConnection] $Connection
                 ,[Parameter(Position=1, Mandatory=$true)] [string] $Query
                 ,[Parameter(Position=2, Mandatory=$false)] [System.Object] $Parameters = @{}
         );
     
     BEGIN
     {
         [System.Data.SqlClient.SqlCommand] $command = $Connection.CreateCommand();
         $command.CommandText = $Query;
         $command.CommandTimeout = $null;
 
         ## Load parameters:
         foreach($key in $Parameters.Keys)
         {
             [void] $command.Parameters.AddWithValue("@$key",  $(if ( $null -eq $Parameters[$key] ) { [DBNull]::Value; } else { $Parameters[$key]; }) );
         }
     }
 
     PROCESS
     {
         $command.ExecuteNonQuery();
     }    
 
     END
     {}
 }
 
 function Sync-SqlData
 {
     <# 
      .SYNOSPIS 
       Sync SQL data.
 
      .DESCRIPTION
       Bulk copy sql data from query results to table.
 
      .PARAMETER Connection
       SQL Server database connection to close.
 
      .PARAMETER Table
       Name of the table to copy to.
 
      .PARAMETER Data
       Data to copy.
 
       .EXAMPLE
       Sync-SqlData -Connection $db -Table 'Accounts_BACKUP' -Data (Get-SqlData -Connection $db -Query 'SELECT id, username, backup_date=CONVERT(DATE,CURRENT_TIMESTAMP) FROM Accounts');
 
       Sync data from Accounts to Accounts_Backup account.
 
     #>
     [CmdletBinding()] 
         param( 
             [Parameter(Position=0, Mandatory=$true)] [System.Data.SqlClient.SqlConnection] $Connection
                 ,[Parameter(Position=1, Mandatory=$true)] [string] $Table
                 ,[Parameter(Position=2, Mandatory=$false)] $Data
         );
     
     BEGIN
     {	
         $bulkCopy = New-Object Data.SqlClient.SqlBulkCopy($Connection);
         $bulkCopy.DestinationTableName = $TableNTableame;
     }
 
     PROCESS
     {
         $bulkCopy.WriteToServer( $Data );
     }
 
     END
     {}
 }
  