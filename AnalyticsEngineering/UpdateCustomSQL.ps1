#Requires -Version 3
Set-StrictMode -Version Latest

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################

## EDIT THIS SECTION
$myTableauServer     = "<your Tableau Server>"
$myTableauServerUser = "<your username>"
$myTableauServerPW   = "<password file>"
$myTableauServerSite = "<Tableau Server site #, if appropriate>"
$myTableauServerSiteName = "<your Tableau Server site Name>"
## DONE EDITING PORTION


Set-Location 'C:\<mk a dir>\TsCustDatasources'
Remove-Item -Path 'C:\<mk a dir>\TsCustDatasources\*' -Recurse -Force -Verbose
Add-Type -AssemblyName System.IO.Compression.Filesystem


#############################################################################
# Taking a master file with a dynamic set of names,companies, etc           # 
# then updating the Tableau Custom SQL, then republishing to Tableau Server # 
# NOTE: The 'C:\<mk a dir>\Names.csv' will be the file that is dynamically  #
# changing.                                                                 # 
#############################################################################

Remove-Item C:\<mk a dir>\ViewSQL\* -Recurse -Verbose -Exclude 'SQLMaster.txt'

Import-CSV 'C:\<mk a dir>\Names.csv' -Delimiter ";" | select -expand Name | select @{n='Company';e={"'$_',"}} | Export-Csv -path 'C:\<mk a dir>\ViewSQL\CustomSQL.csv' -NoTypeInformation -Force -Delimiter ";"
(Import-CSV 'C:\<mk a dir>\ViewSQL\CustomSQL.csv' -Delimiter ";" | select -expand Company -last 1).replace(",","") > 'C:\<mk a dir>\ViewSQL\LastPlan.txt'

$lastplan = gc 'C:\<mk a dir>\ViewSQL\LastPlan.txt'
Import-CSV 'C:\<mk a dir>\ViewSQL\CustomSQL.csv' -Delimiter ";"  | where {$_.Company -notmatch $lastplan} | select -expand Company > 'C:\<mk a dir>\ViewSQL\AllButLastCustomSQL.csv'
gc 'C:\<mk a dir>\ViewSQL\AllButLastCustomSQL.csv', 'C:\<mk a dir>\ViewSQL\LastPlan.txt' > 'C:\<mk a dir>\ViewSQL\FinalSQL.txt'

$sql = (gc C:\<mk a dir>\ViewSQL\FinalSQL.txt)
$query = (gc C:\<mk a dir>\ViewSQL\SQLMaster.txt)
$fullquery = $query | ForEach-Object {$_.replace("x_Replace_X","$($sql)")} ##Future version will replace "x_Replace_X" with a GUID


#####################################################
# Download the datasource we'll need to change      #
# from Tableau Server                               #
#####################################################

$connectionString = 'Driver={PostgreSQL ANSI(x64)};Server=<your server>; Port=<Port>; Database=workgroup; Uid=<user id>; Pwd=<pwd>;'
$query = @"
select 
repository_url as datasource_url
from datasources 
where site_id = $myTableauServerSite and repository_url like '<Tableau Server datasource that needs changing>'
"@

Get-TsExtract -connectionString $connectionString -query $query |`
Select-Object -Skip 1 -Property datasource_url | Export-Csv -Path ./TsDatasources.csv -Delimiter ";" -NoTypeInformation -Force


tabcmd login -s https://$myTableauServer -t $myTableauServerSiteName -u $myTableauServerUser --password-file $myTableauServerPW --no-certcheck
foreach( $i in @(Import-Csv -Path .\TsDatasources.csv -Delimiter ";"|select -expand datasource_url)) {
    tabcmd get "/datasources/$i.tdsx" --filename $i".tdsx"
    Rename-Item $i".tdsx" -NewName $i".zip"
    $pathToZip = "C:\<mk a dir>\TsCustDatasources\$i.zip"
    $targetDir = "C:\<mk a dir>\TsCustDatasources\$i"
    [System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)
}
    
    
#####################################################
# Now update the Tableau xml                        #
# with the new names/companies, etc  and republish  #
# back to Tableau Server to refresh                 #
#####################################################

Remove-Item C:\<mk a dir>\SomeTableauDatasourceFinal\* -Recurse -Verbose 

$k = Get-ChildItem 'C:\<mk a dir>\TsCustDatasources\*\*' -Filter *.tds
$content = "$($k.DirectoryName)\$($k.Name)"
$xmldata = New-Object XML
$xmldata.Load($content)
$newdata = $xmldata.datasource.connection.relation.'#text' = $fullquery | % {$_.Replace("&#13;&#10;"," ") }
$xmldata.Save("C:\<mk a dir>\SomeTableauDatasourceFinal\SomeTableauDatasource.tds") 


remove-item -Path C:\<mk a dir>\TsCustDatasources\SomeTableauDatasource\SomeTableauDatasource.tds
remove-item -Path C:\<mk a dir>\TsCustDatasources\SomeTableauDatasource.zip

copy -Path C:\<mk a dir>\SomeTableauDatasourceFinal\SomeTableauDatasource.tds -dest C:\<mk a dir>\TsCustDatasources\SomeTableauDatasource\SomeTableauDatasource.tds

## rezip the new file -- already loaded assembly above ##
[System.IO.Compression.ZipFile]::CreateFromDirectory("C:\<mk a dir>\TsCustDatasources\SomeTableauDatasource\","C:\<mk a dir>\TsCustDatasources\SomeTableauDatasource.zip")

## convert back to tableau tdsx ##
Rename-Item "C:\<mk a dir>\TsCustDatasources\SomeTableauDatasource.zip" -NewName 'C:\<mk a dir>\TsCustDatasources\SomeTableauDatasource.tdsx' -Force


tabcmd login -s https://$myTableauServer -t $myTableauServerSiteName -u $myTableauServerUser --password-file $myTableauServerPW  --no-certcheck
tabcmd publish --project "Default" "SomeTableauDatasource.tdsx" --name "SomeTableauDatasource" --overwrite --oauth-username $myTableauServerUser --save-oauth --no-certcheck
tabcmd refreshextracts --project "Default" --datasource "SomeTableauDatasource"


tabcmd logout
