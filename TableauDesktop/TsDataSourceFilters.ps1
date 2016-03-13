#Requires -Version 3
#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################

Set-StrictMode -Version Latest
Set-Location 'C:\TsAdmin\TsDatasource'
Remove-Item -Path C:\TsAdmin\TsDatasource\* -Recurse -Force -Verbose -Exclude '*.csv' #need to remove old tds stuff to avoid conflicts
Import-Module 'C:\Users\<your username>\Documents\WindowsPowerShell\Modules\TsGetExtract\TsGetExtract.psm1'

#Load the .Net assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null

## EDIT THIS SECTION
$myTableauServer     = "<your server>"
$myTableauServerUser = "<your username>"
$myTableauServerPW   = "<path to pw file>"
$myTableauServerSite = "1"
## DONE EDITING PORTION


## NOTHING NEEDS CHANGING PAST THIS POINT
switch ("$myTableauServerSite") {
    "3" {$TsPubSite = "<site>";break}
    "1" {$TsPubSite = "<site>";break}
    "2" {$TsPubSite = "<site>";break}
    "6" {$TsPubSite = "<site>";break} 
}

$connectionString = 'Driver={PostgreSQL ANSI(x64)};Server=<your server>; Port=8060; Database=workgroup; Uid=<your Ts Un>; Pwd=<your Ts Pw>;'
$query = @"
select datasource_url, site_id
from _datasources 
where dbclass like '<something>' and site_id = $myTableauServerSite
"@

Get-TsExtract -connectionString $connectionString -query $query |`
Select-Object -Skip 1 -Property datasource_url  | Export-Csv -Path ./TsDatasources.csv -Delimiter ";" -NoTypeInformation -Force

## download the datasource from server

if ($myTableauServerSite -ne 1) {
    tabcmd login -t $TsPubSite -s https://$myTableauServer -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
    foreach( $i in @(Import-Csv -Path .\TsDatasources.csv -Delimiter ";"|select -expand datasource_url)) {
        tabcmd get https://$myTableauServer/t/$TsPubSite/datasources/$i.tdsx --filename $i".tdsx"
        Rename-Item $i".tdsx" -NewName $i".zip"
        $pathToZip = "C:\TsAdmin\TsDatasource\$i.zip"
        $targetDir = "C:\TsAdmin\TsDatasource\$i"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)
        }} else {
        tabcmd login -s https://$myTableauServer -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
        foreach( $i in @(Import-Csv -Path .\TsDatasources.csv -Delimiter ";"|select -expand datasource_url)) {
            tabcmd get https://$myTableauServer/datasources/$i.tdsx --filename $i".tdsx"
            Rename-Item $i".tdsx" -NewName $i".zip"
            $pathToZip = "C:\TsAdmin\TsDatasource\$i.zip"
            $targetDir = "C:\TsAdmin\TsDatasource\$i"
            [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)
        } #end foreach
      } #end else
    
    

## update the xml ##
Remove-Item -Path C:\TsAdmin\TsDatasource\TsDatasourcesGroupFilter.csv -Force -ErrorAction SilentlyContinue
foreach ($k in @(Get-ChildItem 'C:\TsAdmin\TsDatasource\*\*' -Filter *.tds)) {
    $content = "$($k.DirectoryName)\$($k.Name)"
    $xmldata = New-Object XML
    $xmldata.Load($content)
    $xmldataFilter = $xmldata.datasource | select -Property Filter
    if ( ($xmldataFilter.Filter -eq $null) -eq $true ) {out-null} else {
        $DsFilter = $xmldata.datasource.filter | select InnerXml, @{n='Datasource';e={$k}}, @{n='Site';e={$TsPubSite}}
        $DsFilter | Export-Csv -Path C:\TsAdmin\TsDatasource\TsDatasourcesGroupFilter.csv -Delimiter ";" -NoTypeInformation -Append
     } #end else
} #end foreach

tabcmd logout

##LOCAL PROCESSING: Parsing file from above to split out actual fields from datasource filter
foreach($t in @(ipcsv .\TsDatasourcesGroupFilter.csv -Delimiter ";" | select *)) {
[xml]$x = @" 
$($t.InnerXml)
"@
    if ($x.groupfilter.function -eq 'union' -and $x.groupfilter.'ui-enumeration' -eq 'inclusive') {$k = $x.SelectNodes("//groupfilter/*") | select -expand member -ErrorAction SilentlyContinue
            $x.groupfilter | select -Property ui-domain, ui-enumeration, @{n='field';e={$x.groupfilter.groupfilter | `
            select -expand level -Unique}}, @{n='field_member(s)';e={$k -join ","}}, @{n='File';e={$t.Datasource}}, @{n='site';e={$t.Site}},@{n='Date';e={Get-Date}} | `
            Export-Csv -Path C:\TsAdmin\TsDatasource\TsDatasourcesGroupFilterFields.csv -Delimiter ";" -NoTypeInformation -Append
            } #end if 
            
            elseif ($x.groupfilter.'ui-enumeration' -eq 'inclusive') {$x.groupfilter | `
                 select -Property ui-domain, ui-enumeration, @{n='field';e={$_.level}}, @{n='field_member(s)';e={$_.member}},@{n='File';e={$t.Datasource}},@{n='site';e={$t.Site}},@{n='Date';e={Get-Date}} | `
                 Export-Csv -Path C:\TsAdmin\TsDatasource\TsDatasourcesGroupFilterFields.csv -Delimiter ";" -NoTypeInformation -Append
                 } #end elseif 
                 else{
                        $s = $x.SelectNodes("//groupfilter/*") | select -expand member -ErrorAction SilentlyContinue
                        $x.groupfilter | select -Property ui-domain, ui-enumeration, @{n='field';e={$x.groupfilter.groupfilter | `
                        select -expand level -Unique}}, @{n='field_member(s)';e={$s -join ","}}, @{n='File';e={$t.Datasource}}, @{n='site';e={$t.Site}},@{n='Date';e={Get-Date}} | `
                        Export-Csv -Path C:\TsAdmin\TsDatasource\TsDatasourcesGroupFilterFields.csv -Delimiter ";" -NoTypeInformation -Append
                 } #end else
                 
} #end foreach
