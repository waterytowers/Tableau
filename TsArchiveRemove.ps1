#Requires -Version 3
Set-StrictMode -Version Latest

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################

##     Replace '<directory path>' with your own directory. Example: C:\TsGetMetadata
##     This is also the 2nd part after you've downloaded all the workbooks (or whatever wbs / datasources you need)

$today = (get-date).ToString('yyyy-MM-dd')
$myTableauServer     = "<your server>"
$myTableauServerUser = "<your user>"
$myTableauServerPW   = "<your pw/pw file> "
$copyPath            = "<copy path if appropriate>"
$wbExportPath        = "<Path to all the twb and twbx files>"


if (Test-Path '<directory path>') {
    Set-Location '<directory path>'
    } else {
        mkdir '<directory path>' ; Set-Location '<directory path>'
       }

Remove-Item -Path <directory path>\* -Recurse -Force -Verbose -Exclude '*.csv' 
Remove-Item -Path <directory path>\MetadataErrors.txt -Verbose -Force -ErrorAction SilentlyContinue

#Load the .Net assembly
Add-Type -AssemblyName system.io.compression.filesystem


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

ipcsv .\TsDatasourcesTs.csv -Delimiter ";" | where {$_.site_id -eq 1} | select * | Export-Csv <directory path>\TsDatasourcesTsCert.csv -Delimiter ";" -NoTypeInformation -Force
ipcsv .\TsDatasourcesTs.csv -Delimiter ";" | where {$_.site_id -ne 1} | select * | Export-Csv <directory path>\TsDatasourcesTsOther.csv -Delimiter ";" -NoTypeInformation -Force

# Other sites
 foreach( $datasource in @(Import-Csv -Path <directory path>\TsDatasourcesTsOther.csv -Delimiter ";")) {

        tabcmd login -s http://$myTableauServer -t "$($datasource.Site_Name)" -u $myTableauServerUser -p $myTableauServerPW 
        tabcmd get "/datasources/$($datasource.ds_url).tdsx" --filename "$($datasource.ds_url).tdsx"
        
        Rename-Item "$($datasource.ds_url).tdsx" -NewName "$($datasource.ds_url)_$($datasource.Site_Name).zip"
        
        $pathToZip = "<directory path>\$($datasource.ds_url)_$($datasource.Site_Name).zip"
        $targetDir = "<directory path>\$($datasource.ds_url)_$($datasource.Site_Name)"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)

        $k = Get-ChildItem "<directory path>\$($datasource.ds_url)_$($datasource.Site_Name)\*" -Filter *.tds
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
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <directory path>\MetadataErrors.txt }
        
        Try 
        { $customsql = $xmldata.datasource.connection.relation.'#text' }
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <directory path>\MetadataErrors.txt }
        
        Try 
        { $elfile = $xmldata.datasource.connection.filename}
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <directory path>\MetadataErrors.txt }


        $xmldata.SelectNodes('//datasource/column') | select -Property  @{n='Segment';e={'Datasource'}},Caption,DataType,'default-format',Name,Role,Type,@{n='Calculation';e={$_.calculation.formula}}, @{n='DatasourceName';e={$k.DirectoryName -split "\\" | select -Last 1}},@{n='DatasourceCaption';e={$null}},@{n='Site';e={$datasource.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$null}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <directory path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append

}

tabcmd logout

# Default site
 foreach( $datasource in @(Import-Csv -Path <directory path>\TsDatasourcesTsCert.csv -Delimiter ";")) {

        tabcmd login -s http://$myTableauServer -u $myTableauServerUser -p $myTableauServerPW 
        tabcmd get "/datasources/$($datasource.ds_url).tdsx" --filename "$($datasource.ds_url).tdsx"
        
        Rename-Item "$($datasource.ds_url).tdsx" -NewName "$($datasource.ds_url)_$($datasource.Site_Name).zip"
        
        $pathToZip = "<directory path>\$($datasource.ds_url)_$($datasource.Site_Name).zip"
        $targetDir = "<directory path>\$($datasource.ds_url)_$($datasource.Site_Name)"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)

        $k = Get-ChildItem "<directory path>\$($datasource.ds_url)_$($datasource.Site_Name)\*" -Filter *.tds
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
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <directory path>\MetadataErrors.txt }
        
        Try 
        { $customsql = $xmldata.datasource.connection.relation.'#text' }
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <directory path>\MetadataErrors.txt }
        
        Try 
        { $elfile = $xmldata.datasource.connection.filename}
        Catch 
        {"PropertyNotFoundStrict: $($datasource.ds_url)" *>> <directory path>\MetadataErrors.txt }


        $xmldata.SelectNodes('//datasource/column') | select -Property  @{n='Segment';e={'Datasource'}},Caption,DataType,'default-format',Name,Role,Type,@{n='Calculation';e={$_.calculation.formula}}, @{n='DatasourceName';e={$k.DirectoryName -split "\\" | select -Last 1}},@{n='DatasourceCaption';e={$null}},@{n='Site';e={$datasource.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$null}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <directory path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append

}



##############################################
##           Workbooks section              
##############################################


## ##############################################################
## Download the workbook(s) with the, 'TsExportAllWorkbooks' 
## script 1st and then do this parsing from server b/c they'll have 
## more (usually) data sources embedded in the workbook
#################################################################

Set-Location $wbExportPath

foreach($i in @(Get-ChildItem -Path $wbExportPath )) {
    if ($i.extension -eq ".twbx") { 
                #Write-Verbose -Message "This is a twbx" -Verbose
                
                Rename-Item "$($i.name)" -NewName "$($i.basename).zip"
    
                $pathToZip = "$wbExportPath\$($i.basename).zip"
                $targetDir = "$wbExportPath\$($i.basename)"
                [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)
    
                $k = gci -path "$wbExportPath\$($i.basename)\*.twb"
                $content = "$($k.DirectoryName)\$($k.Name)"
                $m  = "$($k.DirectoryName)\New$($k.Name)"
                $x  = "$($k.DirectoryName)\NewNew$($k.Name)"

                gc $content | ForEach-Object {$_.Replace("&#10;","")} | Set-Content -Path $m
                gc $m | ForEach-Object {$_.Replace("&#13;","")} | Set-Content -path $x

                $xmldata = New-Object XML
                $xmldata.Load($x)

                $ErrorActionPreference = "SilentlyContinue"
                foreach( $elItem in @($xmldata.workbook.datasources.datasource)) {
                     $eltable   = $null
                     $customsql = $null
                     $elfile = $null
                     try 
                        {$eltable   = $elItem.connection.relation.table}
                     catch 
                        {"PropertyNotFoundStrict: $($i.wb_url)" *>> <directory path>\MetadataErrors.txt }
     
                     try
                     {$customsql = $elItem.connection.relation.'#text'}
                     catch 
                        {"PropertyNotFoundStrict: $($i.wb_url)" *>> <directory path>\MetadataErrors.txt }
     
                     try
                     {$elfile = $elItem.connection.relation.filename}
                     catch 
                      {"PropertyNotFoundStrict: $($i.wb_url)" *>> <directory path>\MetadataErrors.txt }
     
                     $elitem.column | select -Property @{n='Segment';e={'Workbook'}},Caption,DataType,'default-format',Name,Role,Type, @{n='DatasourceName';e={$elItem.name}},@{n='DatasourceCaption';e={$elItem.caption}}, @{n='Site';e={$i.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$i.name}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql.Replace(" ","|")}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <directory path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append
                 }
                
            } else {
                
                $k = gci -path "$wbExportPath\$($i.name)"
                gc $k | ForEach-Object {$_.Replace("&#10;","")} | Set-Content -Path "$wbExportPath\2_$($i.name)"
                gc "$wbExportPath\2_$($i.name)" | ForEach-Object {$_.Replace("&#13;","")} | Set-Content -path "$wbExportPath\3_$($i.name)"
      
                [xml]$xmldata = (gc "$wbExportPath\3_$($i.name)")

                $ErrorActionPreference = "SilentlyContinue"
                foreach( $elItem in @($xmldata.workbook.datasources.datasource)) {
                     $eltable   = $null
                     $customsql = $null
                     $elfile    = $null
                     try 
                        {$eltable   = $elItem.connection.relation.table}
                     catch 
                        {"PropertyNotFoundStrict: $($i.name)" *>> <directory path>\MetadataErrors.txt }
     
                     try
                     {$customsql = $elItem.connection.relation.'#text'}
                     catch 
                        {"PropertyNotFoundStrict: $($i.Name)" *>> <directory path>\MetadataErrors.txt }
     
                     try
                     {$elfile = $elItem.connection.relation.filename}
                     catch 
                      {"PropertyNotFoundStrict: $($i.name)" *>> <directory path>\MetadataErrors.txt }
     
                     $elitem.column | select -Property @{n='Segment';e={'Workbook'}},Caption,DataType,'default-format',Name,Role,Type, @{n='DatasourceName';e={$elItem.name}},@{n='DatasourceCaption';e={$elItem.caption}}, @{n='Site';e={$i.site_name}},@{n='CurrentDate';e={$today}}, @{n='Source';e={$i.name}}, @{n='Table';e={$eltable}}, @{n='CustomSQL';e={$customsql.Replace(" ","|")}},@{n='DsFileName';e={$elFile}} | Export-Csv -Path <directory path>\TsDatasourcesNameInfo.csv -Delimiter ";" -NoTypeInformation -Append
                 }
            }
               

}

tabcmd logout


# Clean the log output for the workbook/datasource error(s) section
$a = gc -Path <directory path>\MetadataErrors.txt
$b = "Log Detail"  
Set-Content <directory path>\MetadataErrors.txt –value $b, $a[0..($a.Count-1)] 
Import-Csv -Path <directory path>\MetadataErrors.txt | select @{n='Date';e={Get-Date}},*|Export-Csv <directory path>\MetadataErrors.csv -Delimiter ";" -NoTypeInformation -Append
 
copy -Path <directory path>\TsDatasourcesNameInfo.csv -Destination $copyPath -Force

tabcmd logout

$ErrorActionPreference = "Continue"
