#Requires -Version 3

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.Filesystem

$pw      = (gc -Path "<your pw config file>") | ConvertTo-SecureString -AsPlainText -Force
$pwTmp   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw)
$pwFinal = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($pwTmp)

#############################################################   Change me   ###############################################################

$MasterFile          = "<your master file with updates>"
$myTableauServerProd = "<your Tableau Server>"
$myTableauServerUser = "<your Tableau username>"
$myTableauServerSite = "<your Tableau Server site>"
$myDatasource        = "<name of Data source to be replaced>"
$workPath            = "<your work path>"
$UpdatedSQL = @"
<The SQL you want changed>
"@
$DbUsername          = "<your db username>"
$DbPw                = "<your db password>"
 
#############################################################    Done with changes  ########################################################



tabcmd logout

        #############################################
        # Update the tableau server datasource      #
        #############################################

if (Test-Path $workPath) {
    Set-Location $workPath
    } else {
        mkdir $workPath ; Set-Location $workPath
       }
       
Remove-Item ./* -Recurse -Verbose -Force


            #############################################
            # download current datasource               #
            #############################################

tabcmd login -s $myTableauServerProd -site $myTableauServerSite -u $myTableauServerUser -p "$($pwFinal)"
tabcmd get "/datasources/$($myDatasource).tdsx" --filename "$($myDatasource).tdsx"
        
Rename-Item "$($myDatasource).tdsx" -NewName "$($myDatasource).zip"        
$pathToZip = "$workPath\$($myDatasource).zip"
$targetDir = "$workPath\$($myDatasource)"
[System.IO.Compression.ZipFile]::ExtractToDirectory($pathToZip, $targetDir)

            #############################################
            # update custom sql for ds                  #
            #############################################

[xml]$k = gc "$workPath\$($myDatasource)\*" -Filter *.tds
$NewSQL = $k.SelectNodes('//relation') | where {$_.name -like 'custom_sql_query'} | ForEach-Object {$_.'#text' = $UpdatedSQL }    
$k.Save("$workPath\$($myDatasource)\$myDatasource.tds")

## rezip the new file -- already loaded assembly above ##
[System.IO.Compression.ZipFile]::CreateFromDirectory("$workPath\$myDatasource","$workPath\$($myDatasource)_Update.zip")

            #############################################
            # convert back to tdsx and republish        #
            #############################################

Rename-Item "$workPath\$($myDatasource)_Update.zip" -NewName "$workPath\$myDatasource.tdsx" -Force

tabcmd login -s $myTableauServerProd -site $myTableauServerSite -u $myTableauServerUser -p "$($pwFinal)"
tabcmd publish --project "Default" "$myDatasource.tdsx" --name "$myDatasource" --overwrite --db-username "$DbUsername" --db-password "$DbPw" --save-db-password
tabcmd refreshextracts --project "Default" --datasource "$myDatasource"


[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pwTmp)
