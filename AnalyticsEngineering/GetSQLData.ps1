function Get-SQLData {
 [CmdletBinding()]
 param (
    [string]$connectionString,
    [string]$query
    )
    Write-Verbose 'Getting Database Data'
    $connection = New-Object -TypeName System.Data.Odbc.OdbcConnection
    $connection.ConnectionString = $connectionString
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $adapter = New-Object System.Data.Odbc.OdbcDataAdapter $command
    $dataset = New-Object -TypeName System.Data.DataSet
    $adapter.Fill($dataset)
    $dataset.Tables[0]
    $connection.close()
}
