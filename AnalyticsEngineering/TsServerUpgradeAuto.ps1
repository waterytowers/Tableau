#Requires -Version 3.0
function Set-TsUpgrade {
<#
    .SYNOPSIS
        Adds Tableau Server newest version based off of RSS feed

    .DESCRIPTION
        This function will take input from the user in the form of a valid version from Tableau Server
    
    .EXAMPLE
        Set-TsUpgrade -TsVersion 8.3.1
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$TsVersion
    )

<# Omitting this b/c there's a bit of a concern when downloading an exe from a script (no surprise there)
[xml]$a = Invoke-WebRequest -Uri 'http://community.tableausoftware.com/community/feeds/allcontent?community=2014&showProjectContent=false&recursive=false' -Method Get
$b = ($a.feed.entry | select -first 1 | select -expand title) -split "," | select -First 1
$c = $b -replace '[a-zA-Z]',''
$TsVersion = $c.Trim()
#>

############## STEP 1 ##################### 
#  download newest tableau server version # 
###########################################

Remove-Item <your path to server exe>\* -Recurse -Filter *.exe -Force
$Wget = New-Object System.Net.WebClient
$url = "http://downloads.tableausoftware.com/esdalt/$TsVersion/TableauServer-64bit.exe"
$localpath = "<your path to server exe>\TsServer_$TsVersion.exe"
$Wget.DownloadFileAsync($url,$localpath)


do
    {
       "Done with $TsVersion download at: $(Get-Date)" > <file dump confirmation> 
        
    } 
while ($Wget.IsBusy -eq 'False')

#credential stuff to secure email delivery
$credPath = Join-Path (Split-Path $profile) TsUpdateScript.ps1.credential
$credential = Import-Clixml $credPath


############## STEP 2 ####################################
# get latest backup from S3 to restore to new environment#
##########################################################

$TsProdBak = Get-S3Object -BucketName "<your aws bucketname>" -Key "<your aws key>" -Region "<your region>" -StoredCredentials <your StoredCredentials>|where {$_.Size -gt 0}| `
sort -Property LastModified -Descending|select -First 1 -ExpandProperty Key
Copy-S3Object -BucketName "<your aws bucketname>" -Key $TsProdBak -LocalFile "<your file path>\TsProdBak.tsbak" -StoredCredentials <your StoredCredentials>
Start-Sleep -Seconds 420


############ STEP 3 ##################
# uninstall current version          #
###################################### 

tabadmin stop
$uninstall64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_.DisplayName -like "Tableau Server*"} | select -expand Uninstallstring
Start-Process $uninstall64.Replace("`"","") -ArgumentList "/SILENT"
## note, if there is an error it's b/c of the older versions of server that may also be here

############ STEP 4 #####################
# Install of new Tableau Server version #
#########################################

Start-Sleep -Seconds 240
Get-ItemProperty -Path <your path>\TsServer_$TsVersion.exe | select -Property Versioninfo | Format-List > <your path>\TsWin.txt
$TsWindowCaption = Import-Csv -Path <your path>\TsWin.txt -Delimiter ":" -Header a,b,c,d | where {$_.a -like "Product"} | select -expand b
$pinvokes = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr FindWindow(string className, string windowName);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
'@

Add-Type -AssemblyName System.Windows.Forms
Add-Type -MemberDefinition $pinvokes -Name NativeMethods -Namespace MyUtils

Start-Process "<your path>\TsServer_$TsVersion.exe"
Start-Sleep -Seconds 5

# Win 1: Welcome
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim()) 
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5

# Win 2: Destination
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim()) 
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 90

# Sys Verify: part 1
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 5

# Sys Verify: part 2
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 5

# Sys Verify, part 3: hitting next
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim()) 
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5

# radio button to agree, part 1
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 5

# radio button to agree, part 2: up to agree
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim()) 
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{UP}")
Start-Sleep -Seconds 5

# radio button to agree, part 3: tab to go to next
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Seconds 5

#radio button pass
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5

#tab server path: start menu folder
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim()) 
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5

#install 
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 300

#window to config dialog: 'click next to configure tableau server'
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 420

# 1st config window
$hwnd = [MyUtils.NativeMethods]::FindWindow("#32770","Tableau Server Configuration") 
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 60

# Window to confirm install...longest portion
$hwnd = [MyUtils.NativeMethods]::FindWindow("#32770","Tableau Server Configuration")
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 600

## tableau server monitor
$hwnd = [MyUtils.NativeMethods]::FindWindow("TWizardForm","Setup - $TsWindowCaption".Trim())
[MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5

############## STEP 5 ##########
# Restore prod backup to dev   #
################################

gps -Name firefox | kill # change to whatever your default browser is here
$TsStatusBk = tabadmin status
if ($TsStatusBk -like 'Status: STOPPED') {
    start-process "cmd.exe" "/C tabadmin restore <path to backup>\TsProdBak.tsbak" -WindowStyle Hidden `
    -RedirectStandardOutput C:\TableauBackups\ProdBackups\RestoreBackup.txt -RedirectStandardError C:\TableauBackups\ProdBackups\RestoreBackupError.txt
    } 
    else {
            tabadmin stop;Start-Sleep -Seconds 60;start-process "cmd.exe" "/C tabadmin restore <path to backup>\TsProdBak.tsbak" -WindowStyle Hidden `
            -RedirectStandardOutput C:\TableauBackups\ProdBackups\RestoreBackup.txt -RedirectStandardError C:\TableauBackups\ProdBackups\RestoreBackupError.txt
            }

start-sleep -seconds 900

 $TsStatus = tabadmin status
 if ($TsStatus -like 'Status: RUNNING') {Out-Null} 
    else {tabadmin start}

 $Wget.Dispose()
 #gps -Name firefox | kill

$TsVersionAfter = tabadmin version
Start-Sleep -Seconds 10


## Send text upon finish
$SendTo     = "<your provider's text/email info"
Send-MailMessage -Subject 'Tableau Server: Dev Update' -Body "Before: $($TsVersionBefore) `n After: $($TsVersionAfter)" -To $SendTo -From "<sender email>" `
-SmtpServer "<your smtp server>" -Credential $credential -UseSsl

$ErrorActionPreference= 'continue'
