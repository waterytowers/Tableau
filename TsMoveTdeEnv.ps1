#Requires -Version 3
Set-StrictMode -Version Latest

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################

if (Test-Path 'C:\<directory path>\TsMoveToProd') {
    Set-Location 'C:\<directory path>\TsMoveToProd'
    } else {
        mkdir 'C:\<directory path>\TsMoveToProd' ; Set-Location 'C:\<directory path>\TsMoveToProd'
       }

Remove-Item *.tdsx -Verbose -Force


$myTableauDEVServer     = "<your Tableau dev server>"
$myTableauPRODServer    = "<your Tableau dev server>"
$myTableauServerUser 	= "<your Tableau Server user>"
$pw 					= (gc -Path <directory path config pw>) | ConvertTo-SecureString -AsPlainText -Force
$pwTmp 					= [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw)
$pwFinal 				= [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($pwTmp)


## Query Data from Postgres for specific datasource(s) or workbook(s) you want to move from a dev env to a prod env

$connectionString = 'Driver={PostgreSQL ANSI(x64)};Server=<your tableau server>; Port=<your port>; Database=workgroup; Uid=<your tableau username>; Pwd=<your tableau pw>;'
$query = @"
SELECT datasource_url, site_id, s.name as site_name
FROM _datasources 
  JOIN sites s ON s.id = _datasources.site_id
WHERE datasource_url LIKE '<your specific datasource(s) you want to move>'
"@

$MoveDbs = Get-TsExtract -connectionString $connectionString -query $query | Select-Object -Skip 1 -Property datasource_url, site_id, site_name
$TsMoveMe = $MoveDbs.site_name | select -Unique | where {$_ -like '<if multiple sites, you would filter here>'}

## download the datasource or workbook from server

tabcmd login -t $TsMoveMe -s https://$myTableauDEVServer -u $myTableauServerUser -p $pwFinal --no-certcheck
foreach( $i in @($MoveDbs)) {
    tabcmd get "/datasources/$($i.datasource_url).tdsx" --filename "$($i.datasource_url).tdsx"
}

tabcmd logout

## Once you have the datasource(s) exported, now you want to publish them to your prod server 

$MoveDbs = Get-ChildItem -Path 'C:\<directory path>\TsMoveToProd'
tabcmd login -s $myTableauPRODServer -u $myTableauServerUser -p $pwFinal --no-certcheck
foreach( $w in @($MoveDbs)) {
    tabcmd publish --project '<project location>' $($w.name) --overwrite --db-username $myTableauServerUser --db-password $pwFinal
}

tabcmd logout

[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pwTmp)
