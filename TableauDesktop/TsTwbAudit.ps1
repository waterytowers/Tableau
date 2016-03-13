#Requires -Version 5
Set-StrictMode -Version Latest

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################




###########################################################   Change     ##########################################################

# export path
$Export = "<Where you want all this info to go>"


# Getting list of Tableau Workbooks 
$TsServer = '<your server>'
$TsPort = '<your port>'
$TsDb = '<your db>'
$TsUser = '<your Tableau PostgreSQL user>'
$TsPwd = '<your Tableau PostgreSQL pw>'
$myTableauServer = '<your Tableau server>'
$myTableauServerUser = '<your Tableau server user>'
$myTableauServerPW = (Get-Credential).GetNetworkCredential().Password

####################################################### Done with Changes ########################################################





if(!(Test-Path $export)) {
    New-Item -ItemType directory $export -Verbose
    }
        else {Write-Verbose -Message "Directory exists. Will store workbook metadata here: $($Export)" -Verbose
        }


if(!(Test-Path $export\TWB)) {
    New-Item -ItemType directory $export\TWB -Verbose
    }
        else {Write-Verbose -Message "TWB Directory exists." -Verbose
        }
		
if(!(Test-Path $export\TWBArchive)) {
    New-Item -ItemType directory $export\TWBArchive -Verbose
    }
        else {Write-Verbose -Message "TWB Archive Directory exists." -Verbose
        }


Get-ChildItem $Export -Filter *.twbx | Remove-Item -Verbose

# current date
$curdate = (Get-Date).ToString('yyyy-MM-dd')

$connectionString = "Driver={PostgreSQL ANSI(x64)};Server=$($TsServer); Port=$($TsPort); Database=$($TsDb); Uid=$($TsUser); Pwd=$($TsPwd);"
$query = @"
select 
  w.repository_url as wb_url
, w.luid as wb_luid
, w.size as wb_size
, w.site_id as Site_ID
, data_engine_extracts 
, s.name as Site_Name
from workbooks w 
  join sites s on s.id = w.site_id
"@

Get-SQLData -connectionString $connectionString -query $query |`
Select-Object -Skip 1 -Property wb_url, wb_luid,Site_ID, data_engine_extracts,Site_Name , @{n='sizeKB';e={([decimal]::round($_.wb_size/1KB))}} | Export-Csv -Path "$Export\TsDatasources.csv" -Delimiter ";" -NoTypeInformation -Force

ipcsv "$Export\TsDatasources.csv" -Delimiter ";" | where {$_.site_id -eq 1} | select * | Export-Csv "$Export\TsDatasourcesCert.csv" -Delimiter ";" -NoTypeInformation -Force
ipcsv "$Export\TsDatasources.csv" -Delimiter ";" | where {$_.site_id -ne 1} | select * | Export-Csv "$Export\TsDatasourcesOther.csv" -Delimiter ";" -NoTypeInformation -Force



# Download the Server [DEFAULT SITE] workbooks minus the extracts
foreach($i in @(Import-Csv -Path "$Export\TsDatasourcesCert.csv" -Delimiter ";" )) {
    if($i.data_engine_extracts -eq 1) {
        tabcmd login -s "https://$myTableauServer" -u "$($myTableauServerUser)" -p "`"$myTableauServerPW`"" --no-certcheck
        tabcmd get "/workbooks/$($i.wb_url).twbx?no_extract=1" --filename "$Export\$($i.wb_url).twbx"
    } else {
            tabcmd login -s "https://$myTableauServer" -u "$($myTableauServerUser)" -p "`"$myTableauServerPW`"" --no-certcheck
            tabcmd get "/workbooks/$($i.wb_url).twb" --filename "$Export\$($i.wb_url).twb"
        }
}


Get-ChildItem $Export -Filter *.twb | move -Destination "$Export\TWB"


# Rename / Unzip 
foreach($twbx in @(Get-ChildItem $Export -Filter *.twbx)) {
 Rename-Item "$export\$twbx" -NewName "$export\$($twbx.basename).zip" -Force ; Expand-Archive -Path "$export\$($twbx.basename).zip" -Force -DestinationPath "C:\Users\$env:username\Dropbox (Pluralsight)\mike-roberts\TableauServerMeta\TWB" 
}

# getting rid of everything BUT the twb 
Remove-Item -Path "$Export\TWB\*" -Recurse -Exclude *.twb -Force -Verbose

# don't need the zip files anymore
Get-ChildItem $Export -Filter *.zip | Remove-Item -Verbose


# extracting the workbook info
$ErrorActionPreference = 'SilentlyContinue'

foreach($TsWorkbook in @(Get-ChildItem -Path "$Export/TWB" -Filter *.twb)) {

[xml]$wb = Get-Content "$Export/TWB/$TsWorkbook"

foreach($ds in @($wb.workbook.datasources.datasource) ) {
     $dscaption = $null ; $dsname = $null
     $dscaption = $ds.caption; $dsname = $ds.name; $ds.column | select `
     @{n='FileName';e={$($TsWorkbook.name)}}, `
     @{n='DsType';e={'Column'}}, `
     @{n='DsCaption';e={$dscaption}}, `
     @{n='DsName';e={$dsname}}, `
     caption, `
     datatype, `
     name, `
     role, `
     type, `
     @{n='calc';e={($_.calculation.formula) -replace "`t|`n|`r"," "}}, `
     @{n='aggregation';e={$null}}, `
     @{n='class';e={$null}}, `
     @{n='contains-null';e={$null}}, `
     @{n='family';e={$null}}, `
     @{n='layered';e={$null}}, `
     @{n='local-name';e={$null}}, `
     @{n='local-type';e={$null}}, `
     @{n='ordinal';e={$null}}, `
     @{n='parent-name';e={$null}}, `
     @{n='remote-alias';e={$null}}, `
     @{n='remote-name';e={$null}}, `
     @{n='remote-type';e={$null}}, `
     @{n='ConnectionInfo';e={$null}}, `
     @{n='ExtractDate';e={$curdate}} | Export-Csv "$Export\TsMetaDataInfo.csv" -NoTypeInformation -Delimiter ";" -Append -Force 
     }  # corresponds to what's in the workbook


foreach($ds in @($wb.workbook.datasources.datasource) ) {
     $dscaption = $null ; $dsname = $null
     $dscaption = $ds.caption; $dsname = $ds.name; `
     $ds.connection.'metadata-records'.'metadata-record' | select  `
     @{n='FileName';e={$($TsWorkbook.name)}}, `
     @{n='DsType';e={'Metadata_records'}}, `
     @{n='DsCaption';e={$dscaption}}, `
     @{n='DsName';e={$dsname}}, `
     @{n='caption';e={$null}}, `
     @{n='datatype';e={$null}}, `
     @{n='name';e={$null}}, `
     @{n='role';e={$null}}, `
     @{n='type';e={$null}}, `
     @{n='calc';e={$null}}, `
      aggregation, `
      class, `
      contains-null, `
      family, `
      layered, `
      local-name, `
      local-type, `
      ordinal, `
      parent-name, `
      remote-alias, `
      remote-name, `
      remote-type, `
      @{n='ConnectionInfo';e={$null}}, `
      @{n='ExtractDate';e={$curdate}} | Export-Csv "$Export\TsMetaDataInfo.csv"   -NoTypeInformation -Delimiter ";" -Append -Force 
   }
      # corresponds to 'remote' metadata



foreach($connection in @($wb.workbook.datasources.datasource) ) {
     $dscaption = $null ; $dsname = $null
     $dscaption = $connection.caption; $dsname = $connection.name ; `
     $connection.connection.relation | select `
     @{n='FileName';e={$($TsWorkbook.name)}}, `
     @{n='DsType';e={'DataSourceInfo'}}, `
     @{n='DsCaption';e={$dscaption}}, `
     @{n='DsName';e={$dsname}}, `
     @{n='caption';e={$null}}, `
     @{n='datatype';e={$null}}, `
     @{n='name';e={$null}}, `
     @{n='role';e={$null}}, `
     @{n='type';e={$null}}, `
     @{n='calc';e={$null}}, `
     @{n='aggregation';e={$null}}, `
     @{n='class';e={$null}}, `
     @{n='contains-null';e={$null}}, `
     @{n='family';e={$null}}, `
     @{n='layered';e={$null}}, `
     @{n='local-name';e={$null}}, `
     @{n='local-type';e={$null}}, `
     @{n='ordinal';e={$null}}, `
     @{n='parent-name';e={$null}}, `
     @{n='remote-alias';e={$null}}, `
     @{n='remote-name';e={$null}}, `
     @{n='remote-type';e={$null}}, `
     @{n='ConnectionInfo';e={($_.OuterXml) -replace "`t|`n|`r"," "}}, `
     @{n='ExtractDate';e={$curdate}} | Export-Csv "$Export\TsMetaDataInfo.csv"  -NoTypeInformation -Delimiter ";" -Append -Force 
     } 
    

}

tabcmd logout

# need to archive old twbs
Get-ChildItem -Path "$export\TWB" -filter *.twb  | move -Destination "$export\TWBArchive" -Force

$ErrorActionPreference = 'Continue'

# Do a diff on content of file and directory to ensure we parsed all the TWBs
$FileInfo = Import-Csv -Path $export\TsMetaDataInfo.csv -Delimiter ";" | where {$_.ExtractDate -eq $curdate} | group -Property FileName -NoElement | select -expand name -Unique
$FileList = Get-ChildItem -Path $export\TWBArchive | where {($_.LastWriteTime).ToString('yyyy-MM-dd') -eq $curdate} | select -expand name 


$TsMetaDataInfo = @{
    'Count of Equal' = (Compare-Object -ReferenceObject $FileInfo -DifferenceObject $FileList -IncludeEqual | where {$_.SideIndicator -eq '=='} | select -expand InputObject).Count ;
    'Count of Diff' = (Compare-Object -ReferenceObject $FileInfo -DifferenceObject $FileList -IncludeEqual | where {$_.SideIndicator -ne '=='} | select -expand InputObject).Count
    }

$TsMetaDataInfo

foreach($err in $Error) {
    if ($err -eq $null) {
        $err | select @{n='ErrorDate';e={$(Get-Date)}},@{n='ErrorMessage';e={$null}} | Export-Csv $export\TsMetaDataErrorInfo.csv -Delimiter ";" -NoTypeInformation -Append
        } else {
           $err | select  @{n='ErrorDate';e={$(Get-Date)}},@{n='ErrorMessage';e={($err.Exception.Message)  -replace "`t|`n|`r"," "}} | Export-Csv $export\TsMetaDataErrorInfo.csv -Delimiter ";" -NoTypeInformation -Append
        } 
    } 

