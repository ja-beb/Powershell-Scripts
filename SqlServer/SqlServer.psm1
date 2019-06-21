<#
    Author_____: Sean Bourg <sean.bourg@gmail.com>
    Date_______: 2018-08-01
    Description: Collection of functions for working with an MS SQL Server database.    
 #>

function Open-SqlConnection {
    <# 
      .SYNOPSIS
       Open MS SQL Server connection database file.
 
      .DESCRIPTION
       Open MS SQL Server connection database file. 
 
      .PARAMETER Server
      The server that the database resides on.
 
      .PARAMETER Database
      The name of the database to open once connected.
      
      .PARAMETER Credentials
      Credentials used to connect to the database if different from account executing the current script.
 
      .PARAMETER Timeout
      The number of sections until connection times out.
 
      .OUTPUTS
      System.Data.SqlClient.SqlConnection
 
      .EXAMPLE
      $database = Open-SqlConnection -Server 'SqlServer' -Database 'MyDatabase';
 
      .EXAMPLE
      $credentials = Get-Credential -UserName 'website_user' -Message "Provide password for database"
      $db = Open-SqlConnection -Server 'localhost' -Database 'example-website' -Credentials $credentials;

     #> 
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory = $true)] [string] $Server,
        [Parameter(Mandatory = $true)] [string] $Database,
        [Parameter(Mandatory = $false)] [System.Management.Automation.PSCredential] $Credentials = $null,
        [Parameter(Mandatory = $false)] [int] $Timeout = 15 
    );
 
    $security = if ([string]::IsNullOrEmpty($PSCredential) ) { 
        'Integrated Security=True';
    }
    else {
        'User ID={0};Password={1}' -f $Credentials.GetNetworkCredential().UserName, $Credentials.GetNetworkCredential().Password;        
    } 

    $connectionString = 'Server={0};Database={1};Connect Timeout={2};{3}' -f $Server, $Database, $Timeout, $security;
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString);
    $connection.Open();
    return $connection;
}
 
function Get-SqlData {
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
        [Parameter(Mandatory = $true)] [System.Data.SqlClient.SqlConnection] $Connection,
        [Parameter(Mandatory = $true)] [string] $Query,
        [Parameter(Mandatory = $false)] [hashtable] $Parameters = @{ }
    );
     
    ## Create query:
    [System.Data.SqlClient.SqlCommand] $command = $Connection.CreateCommand();
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
 
function Invoke-SqlQuery {  
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
        [Parameter(Mandatory = $true)] [System.Data.SqlClient.SqlConnection] $Connection,
        [Parameter(Mandatory = $true)] [string] $Query,
        [Parameter(Mandatory = $false)] [hashtable] $Parameters = @{ }
    );
     
    [System.Data.SqlClient.SqlCommand] $command = $Connection.CreateCommand();
    $command.CommandText = $Query;
    $command.CommandTimeout = $null;
 
    ## Load parameters:
    foreach ($key in $Parameters.Keys) {
        [void] $command.Parameters.AddWithValue("@$key", $(if ( $null -eq $Parameters[$key] ) { [DBNull]::Value; } else { $Parameters[$key]; }) );
    }
    $command.ExecuteNonQuery();
}
 
function Sync-SqlData {
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
        [Parameter(Mandatory = $true)] [System.Data.SqlClient.SqlConnection] $Connection,
        [Parameter(Mandatory = $true)] [string] $Table,
        [Parameter(Mandatory = $false)] [System.Data.DataTable] $Data
    );
     
    $bulkCopy = New-Object Data.SqlClient.SqlBulkCopy($Connection);
    $bulkCopy.DestinationTableName = $TableNTableame;
    $bulkCopy.WriteToServer( $Data );
}
  
