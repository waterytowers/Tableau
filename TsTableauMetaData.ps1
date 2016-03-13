#Requires -Version 3
Set-StrictMode -Version Latest

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################

#NOTE: Replace '<some path>' with your own directory. Example: C:\TsGetMetadata

if (Test-Path '<some path>') {
    Set-Location '<some path>'
    } else {
        mkdir '<some path>' ; Set-Location '<some path>'
       }
Remove-Item -Path <some path>\* -Recurse -Force -Verbose -Exclude '*.csv' #need to remove old tds stuff to avoid conflicts
Remove-Item -Path <some path>\MetadataErrors.txt -Verbose -Force -ErrorAction SilentlyContinue

#Load the .Net assembly
Add-Type -AssemblyName system.io.compression.filesystem


$today = (get-date).ToString('yyyy-MM-dd')
$myTableauServer     = "<your tableau server>"
$myTableauServerUser = "<>"
$myTableauServerPW   = "<pw file>"
$copyPath            = "<if necessary, a path to copy to>"



########################################################
##           Datasources section                      ##
## NOTE: workbooks and datasources will be DIFFERENT  ##
## so datasources (Tableau datasources) will need to  ##
## be separate for both Default and Other sites       ##
##                                                    ##
########################################################

$TsconnectionString = 'Driver={PostgreSQL ANSI(x64)};Server=<>; Port=<>; Database=<>; Uid=<>; Pwd=<>;'
$Tsquery = @"
select 
  d.repository_url as ds_url
, d.luid as ds_luid
, d.size as ds_size
, d.site_id as Site_ID
, s.name as Site_Name
from datasources d 
  join sites s on s.id = d.site_id
"@

Get-TsExtract -connectionString $TsconnectionString -query $Tsquery |`
Select-Object -Skip 1 -Property ds_url,ds_luid,Site_ID,Site_Name,@{n='sizeKB';e={([decimal]::round($_.ds_size/1KB))}} | Export-Csv -Path ./TsDatasourcesTs.csv -Delimiter ";" -NoTypeInformation -Force

ipcsv .\TsDatasourcesTs.csv -Delimiter ";" | where {$_.site_id -eq 1} | select * | Export-Csv <some path>\TsDatasourcesTsCert.csv -Delimiter ";" -NoTypeInformation -Force
ipcsv .\TsDatasourcesTs.csv -Delimiter ";" | where {$_.site_id -ne 1} | select * | Export-Csv <some path>\TsDatasourcesTsOther.csv -Delimiter ";" -NoTypeInformation -Force

# Other sites
 foreach( $datasource in @(Import-Csv -Path <some path>\TsDatasourcesTsOther.csv -Delimiter ";")) {

        tabcmd login -s https://$myTableauServer -t $datasource.Site_Name -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
        tabcmd get "/datasources/$($datasource.ds_url).tdsx" --filename "$($datasource.ds_url).tdsx"
        Rename-Item "$($datasource.ds_url).tdsx" -NewName "$($datasource.ds_url).zip"
        $pathToZip = "<some path>\$($datasource.ds_url).zip"
        $targetDir = "<some path>\$($datasource.ds_url)"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)

        $k = Get-ChildItem "<some path>\$($datasource.ds_url)\*" -Filter *.tds
        $content = "$($k.DirectoryName)\$($k.Name)"
        $m  = "$($k.DirectoryName)\New$($k.Name)"
        $x  = "$($k.DirectoryName)\NewNew$($k.Name)"
        gc $content | ForEach-Object {$_.Replace("&#10;","")} | Set-Content -Path $m
        gc $m | ForEach-Object {$_.Replace("&#13;","")} | Set-Content -path $x
        $xmldata = New-Object XML
        $xmldata.Load($x)

        $ErrorActionPreference = "Stop"
    
        $eltable   = $null
        $customsql = $null
        $elfile = $null

        Try 
        {$eltable = $xmldata.datasource.connection.relation.table} 
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <some path>\MetadataErrors.txt }
        
        Try 
        { $customsql = $xmldata.datasource.connection.relation.'#text' }
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <some path>\MetadataErrors.txt }
        
        Try 
        { $elfile = $xmldata.datasource.connection.filename}
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <some path>\MetadataErrors.txt }


        $xmldata.SelectNodes('//datasource/column') | select -Property  @{n='Segment';e={'Datasource'}},Caption,DataType,'default-format',Name,Role,Type,@{n='Calculation';e={$_.calculation.formula}}, @{n='DatasourceName';e={$k.DirectoryName -split "\\" | select -Last 1}},@{n='DatasourceCaption';e={$null}},@{n='Site';e={$datasource.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$null}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <some path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append

}

tabcmd logout

# Default site
 foreach( $datasource in @(Import-Csv -Path <some path>\TsDatasourcesTsCert.csv -Delimiter ";")) {

        tabcmd login -s https://$myTableauServer -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
        tabcmd get "/datasources/$($datasource.ds_url).tdsx" --filename "$($datasource.ds_url).tdsx"
        Rename-Item "$($datasource.ds_url).tdsx" -NewName "$($datasource.ds_url).zip"
        $pathToZip = "<some path>\$($datasource.ds_url).zip"
        $targetDir = "<some path>\$($datasource.ds_url)"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)

        $k = Get-ChildItem "<some path>\$($datasource.ds_url)\*" -Filter *.tds
        $content = "$($k.DirectoryName)\$($k.Name)"
        $m  = "$($k.DirectoryName)\New$($k.Name)"
        $x  = "$($k.DirectoryName)\NewNew$($k.Name)"
        gc $content | ForEach-Object {$_.Replace("&#10;","")} | Set-Content -Path $m
        gc $m | ForEach-Object {$_.Replace("&#13;","")} | Set-Content -path $x
        $xmldata = New-Object XML
        $xmldata.Load($x)

        $ErrorActionPreference = "Stop"
    
        $eltable   = $null
        $customsql = $null
        $elfile = $null

        Try 
        {$eltable = $xmldata.datasource.connection.relation.table} 
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <some path>\MetadataErrors.txt }
        
        Try 
        { $customsql = $xmldata.datasource.connection.relation.'#text' }
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <some path>\MetadataErrors.txt }
        
        Try 
        { $elfile = $xmldata.datasource.connection.filename}
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <some path>\MetadataErrors.txt }


        $xmldata.SelectNodes('//datasource/column') | select -Property  @{n='Segment';e={'Datasource'}},Caption,DataType,'default-format',Name,Role,Type,@{n='Calculation';e={$_.calculation.formula}}, @{n='DatasourceName';e={$k.DirectoryName -split "\\" | select -Last 1}},@{n='DatasourceCaption';e={$null}},@{n='Site';e={$datasource.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$null}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <some path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append

}

##############################################
##           Workbooks section              ##
##############################################


$connectionString = 'Driver={PostgreSQL ANSI(x64)};Server=<>; Port=<>; Database=<>; Uid=<>; Pwd=<>;'
$query = @"
select 
  w.repository_url as wb_url
, w.luid as wb_luid
, w.size as wb_size
, w.site_id as Site_ID
, s.name as Site_Name
from workbooks w 
  join sites s on s.id = w.site_id
"@

Get-TsExtract -connectionString $connectionString -query $query |`
Select-Object -Skip 1 -Property wb_url, wb_luid,Site_ID, Site_Name , @{n='sizeKB';e={([decimal]::round($_.wb_size/1KB))}} | Export-Csv -Path ./TsDatasources.csv -Delimiter ";" -NoTypeInformation -Force

ipcsv .\TsDatasources.csv -Delimiter ";" | where {$_.site_id -eq 1} | select * | Export-Csv <some path>\TsDatasourcesCert.csv -Delimiter ";" -NoTypeInformation -Force
ipcsv .\TsDatasources.csv -Delimiter ";" | where {$_.site_id -ne 1} | select * | Export-Csv <some path>\TsDatasourcesOther.csv -Delimiter ";" -NoTypeInformation -Force

## download the Other (non-Default) workbook(s) from server b/c they'll have more (usually) data sources embedded in the workbook
foreach($i in @(Import-Csv -Path <some path>\TsDatasourcesOther.csv -Delimiter ";" )) {

    tabcmd login -t "$($i.site_name)" -s https://$myTableauServer -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
    tabcmd get "/workbooks/$($i.wb_url).twbx" --filename "$($i.wb_url).twbx"
    Rename-Item "$($i.wb_url).twbx" -NewName "$($i.wb_url).zip"
    $pathToZip = "<some path>\$($i.wb_url).zip"
    $targetDir = "<some path>\$($i.wb_url)"
    [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)
    
    $k = gci -path "<some path>\$($i.wb_url)\*.twb"
    $content = "$($k.DirectoryName)\$($k.Name)"
    $m  = "$($k.DirectoryName)\New$($k.Name)"
    $x  = "$($k.DirectoryName)\NewNew$($k.Name)"

    gc $content | ForEach-Object {$_.Replace("&#10;","")} | Set-Content -Path $m
    gc $m | ForEach-Object {$_.Replace("&#13;","")} | Set-Content -path $x

    $xmldata = New-Object XML
    $xmldata.Load($x)

    $ErrorActionPreference = "Stop"
    foreach( $elItem in @($xmldata.workbook.datasources.datasource)) {
     $eltable   = $null
     $customsql = $null
     $elfile = $null
     try 
        {$eltable   = $elItem.connection.relation.table}
     catch 
        {"PropertyNotFoundStrict: $($i.wb_url)" *>> <some path>\MetadataErrors.txt }
     
     try
     {$customsql = $elItem.connection.relation.'#text'}
     catch 
        {"PropertyNotFoundStrict: $($i.wb_url)" *>> <some path>\MetadataErrors.txt }
     
     try
     {$elfile = $elItem.connection.relation.filename}
     catch 
      {"PropertyNotFoundStrict: $($i.wb_url)" *>> <some path>\MetadataErrors.txt }
     
     $elitem.column | select -Property @{n='Segment';e={'Workbook'}},Caption,DataType,'default-format',Name,Role,Type,@{n='Calculation';e={$_.calculation.formula}}, @{n='DatasourceName';e={$elItem.name}},@{n='DatasourceCaption';e={$elItem.caption}}, @{n='Site';e={$i.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$k.DirectoryName -split "\\" | select -Last 1}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <some path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append
    }

}

tabcmd logout

## download the Default workbook(s) from server b/c they'll have more (usually) data sources embedded in the workbook
foreach($i in @(Import-Csv -Path <some path>\TsDatasourcesCert.csv -Delimiter ";" )) {

    tabcmd login -s https://$myTableauServer -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
    tabcmd get "/workbooks/$($i.wb_url).twbx" --filename "$($i.wb_url).twbx"
    Rename-Item "$($i.wb_url).twbx" -NewName "$($i.wb_url).zip"
    $pathToZip = "<some path>\$($i.wb_url).zip"
    $targetDir = "<some path>\$($i.wb_url)"
    [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)
    
    $k = gci -path "<some path>\$($i.wb_url)\*.twb"
    $content = "$($k.DirectoryName)\$($k.Name)"
    $m  = "$($k.DirectoryName)\New$($k.Name)"
    $x  = "$($k.DirectoryName)\NewNew$($k.Name)"

    gc $content | ForEach-Object {$_.Replace("&#10;","")} | Set-Content -Path $m
    gc $m | ForEach-Object {$_.Replace("&#13;","")} | Set-Content -path $x

    $xmldata = New-Object XML
    $xmldata.Load($x)

    $ErrorActionPreference = "Stop"
    foreach( $elItem in @($xmldata.workbook.datasources.datasource)) {
     $eltable   = $null
     $customsql = $null
     $elfile = $null
     try 
        {$eltable   = $elItem.connection.relation.table}
     catch 
        {"PropertyNotFoundStrict: $($i.wb_url)" *>> <some path>\MetadataErrors.txt }
     
     try
     {$customsql = $elItem.connection.relation.'#text'}
     catch 
        {"PropertyNotFoundStrict: $($i.wb_url)" *>> <some path>\MetadataErrors.txt }
     
     try
     {$elfile = $elItem.connection.relation.filename}
     catch 
      {"PropertyNotFoundStrict: $($i.wb_url)" *>> <some path>\MetadataErrors.txt }
     
     $elitem.column | select -Property @{n='Segment';e={'Workbook'}},Caption,DataType,'default-format',Name,Role,Type,@{n='Calculation';e={$_.calculation.formula}}, @{n='DatasourceName';e={$elItem.name}},@{n='DatasourceCaption';e={$elItem.caption}}, @{n='Site';e={$i.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$k.DirectoryName -split "\\" | select -Last 1}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <some path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append
    }

}

# Clean the log output for the workbook/datasource error(s) section
$a = gc -Path <some path>\MetadataErrors.txt
$b = "Log Detail"  
Set-Content <some path>\MetadataErrors.txt –value $b, $a[0..($a.Count-1)] 
Import-Csv -Path <some path>\MetadataErrors.txt | select @{n='Date';e={Get-Date}},*|Export-Csv <some path>\MetadataErrors.csv -Delimiter ";" -NoTypeInformation -Append
 
copy -Path <some path>\TsDatasourcesNameInfo.csv -Destination $copyPath -Force

tabcmd logout

$ErrorActionPreference = "Continue"
