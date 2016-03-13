#Requires -Version 3

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################

Set-Location "C:\TsWorkbookPrep"             ## set to your working directory

#Load the .Net assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null

## Edit for your workbook(s)  ##
$myTableauServer     = "<your server>"
$myTableauServerUser = "<your username>"
$myTableauServerPW   = "<your server pw>"
$TsFileExport        = "<your file to export from server>"
$TsFileExportReZip   = "<your rezipped file>"                 ## use this for the re-zip
$SetTsCalc           = "<your calculated field>"              ## name of calculated field to change

$tsCalc = @"                                  
SUM([Sales]) / 100
"@                                                            ## new calculation/formula


############################################
## The section below can remain unchanged  #
## as long as variables above are updated  #
############################################


## download the workbook from server
tabcmd login -s https://$myTableauServer -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
tabcmd get https://$myTableauServer/workbooks/$TsFileExport.twbx --filename $TsFileExport".twbx"
#tabcmd logout

## Rename and extract
Rename-Item $TsFileExport".twbx" -NewName $TsFileExport".zip"
$pathToZip = "C:\TsWorkbookPrep\$TsFileExport.zip"
$targetDir = "C:\TsWorkbookPrep\$TsFileExport"

#Unzip the file
[System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)


## parsing workbook to confirm caption ##
<#
$object = [xml](gc C:\TsWorkbookPrep\$TsFileExport\$TsFileExport.twb)
cls; $object | select-xml "//datasources/datasource/column" | select -expand node | `
 % {$a=$_ | select -Expand calculation -ea 0 | select -Property formula;$_} | `
 where {$_.caption -like 'MyTestField'} | select name, caption, @{n='calc';e={$a | select -expand formula}}
 #>


## update the xml ##
$content = "C:\TsWorkbookPrep\$TsFileExport\$TsFileExport.twb"
$xmldata = New-Object XML
$xmldata.Load($content)

#testing new field name
$newdata = $xmldata.SelectNodes('//workbook/datasources/datasource/column') | where {$_.Caption -like "$SetTsCalc"}
$newdata | ForEach-Object {$_.calculation.formula = $tsCalc }    
$xmldata.Save("C:\TsWorkbookPrep\$TsFileExport\$TsFileExport.twb")

## rezip the new file -- already loaded assembly above ##
[System.IO.Compression.ZipFile]::CreateFromDirectory("C:\TsWorkbookPrep\$TsFileExport","C:\TsWorkbookPrep\$TsFileExportReZip.zip")

## convert back to tableau twbx ##
Rename-Item $TsFileExportReZip".zip" -NewName $TsFileExport".twbx" -Force

## login to server and republish wb ##
tabcmd login -s https://$myTableauServer -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
tabcmd delete --workbook $TsFileExport ## deleting old version before we publish new one
tabcmd publish $TsFileExport".twbx" --no-certcheck
tabcmd logout


## Notes: might add PS Data Team script to remove older workbook and archive it ##
