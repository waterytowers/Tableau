#Requires -Version 3.0
function Set-TsUsers {
<#
    .SYNOPSIS
        Adds Tableau Server users via local authentication and emails username and password to user(s). 

    .DESCRIPTION
        This function will take input from the user in the form of (1) a valid site and (2) a comma-delimited list of users to add to 
        their Tableau Server. NOTE: The '""' in the example is your default site which in most cases is the first site you set up on
        Server. 
    
    .EXAMPLE
        Set-TsUsers -sites '""', 'uncertified' -NewUsers "test-test@email.com"
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('""', 'uncertified')]
        [string[]]$sites,

        [Parameter(Mandatory=$true)]
        [string[]]$NewUsers
    )

Set-Location C:\TsAdmin
tabcmd logout # clearing old sessions

Remove-Item TsServerUserLogging.txt;New-Item TsServerUserLogging.txt -ItemType File -Force

## Password characters for tmp password
$ascii=$null; For ($a=48;$a –le 122;$a++) {$ascii+=,[char][byte]$a }

## Creds for Tableau
$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList "<your username>", (Get-Content C:\TsAdmin\PsOutput\<your pw file>.txt | ConvertTo-SecureString)

## New Tableau users to add, might make this a shared sheet to simply import and parse
$NewUsers -split ',' > TsRemoveUsers.csv

## Remove first from Tableau to reset pw (amongst other things)
tabcmd login -s <your tab server> -u "<your username>" --password-file C:\TsAdmin\TsGuide\<your pw file>.txt --no-certcheck
tabcmd deleteusers "TsRemoveUsers.csv" --no-certcheck *>> C:\TsAdmin\TsServerUserLogging.txt

## clearing old stuff
Remove-Item TsSiteUsers.csv -Force
New-Item TsSiteUsers.csv -ItemType File

foreach($a in $NewUsers)
    {
       $a.split(',') | Add-Content -Path TsSiteUsers.csv
    }

# adding user per the createusers argument in the correct order per TS Admin guide
Import-Csv -Path TsSiteUsers.csv -Header 'UserID'|select *, @{n='pw';e={Get-TempPassword -length 10 -sourcedata $ascii}}, @{n='fn';e={$_.UserID}}, `
@{n='license';e={'interactor'}},@{n='admin';e={'none'}}|Export-Csv TsSiteUsers_tmp.csv -NoTypeInformation 

#remove header for add to tableau server
$tsUsers = ((Get-Content -Path TsSiteUsers_tmp.csv).Replace("`"",""))
$tsUsers[1..($tsUsers.count-1)] > C:\TsAdmin\TsSiteUsers_Final.csv

foreach ($site in $sites) {
        if ($site -like 'uncertified') {
            tabcmd login -s <your tab server> -t $site -u "<your username>" --password-file C:\TsAdmin\TsGuide\<your encrypted pw>.txt --no-certcheck
            tabcmd createsiteusers "c:\TsAdmin\TsSiteUsers_Final.csv" --publisher --no-certcheck *>> C:\TsAdmin\TsServerUserLogging.txt
        } else
          {
            tabcmd login -s <your tab server> -t $site -u "<your username>" --password-file C:\TsAdmin\TsGuide\<your encrypted pw>.txt --no-certcheck
            tabcmd createsiteusers "c:\TsAdmin\TsSiteUsers_Final.csv" --no-publisher --no-certcheck *>> C:\TsAdmin\TsServerUserLogging.txt
          }
  }

# clean the log output
$a = (Get-Content C:\TsAdmin\TsServerUserLogging.txt).Replace(","," ")
$b = "Log Detail"  
Set-Content C:\TsAdmin\TsServerUserLogging.txt –value $b, $a[0..($a.Count-1)] 
Import-Csv -Path C:\TsAdmin\TsServerUserLogging.txt | select @{n='Date';e={Get-Date}},*|Export-Csv C:\TsAdmin\TsAddRemoveEmps.csv -Delimiter ";" -NoTypeInformation -Append 


foreach ($user in @(Import-Csv .\TsSiteUsers_tmp.csv -Delimiter ",")) {
    $userCp     = $($user.UserID).Replace(".","%252E")
    $changeTsPw = "<your tab server>/users/$($userCp)"
    $body       = "Login site: <your tab server> `n username: $($user.UserID) `n pw: $($user.pw) `n This password is temporary and will expire in 24 hours `n Please change password here: $($changeTsPw) "
    $SendTo     = $($user.UserID)
    Send-MailMessage -Subject 'Your new < come creative name here > Account' -Body $body -To $SendTo -From "<your username>" -Cc "<your username>" `
    -SmtpServer "<your smtp server>" -Credential $credentials -UseSsl
    }

tabcmd logout
Set-Location $HOME

} #end function


Set-TsUsers -sites '""', 'uncertified' -NewUsers "test-test@email.com"
