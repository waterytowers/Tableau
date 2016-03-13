#Requires -Version 3
Set-StrictMode -Version Latest 

#####################################################
# Developed by Mike Roberts @ Pluralsight Data Team #
# email: mike-roberts@pluralsight.com               #
#####################################################


$Kenobi      = (get-content "<your config file>") -join "`n" | convertfrom-stringdata
$Slack       = "$($Kenobi.Fett)" #your Slack API token 
$workingDir  = "<your working directory>"
$archiveDir  = "<your archive dir>" #location for logging resets
$psAddDir    = "<your file for the tableau server users that need reset>"
$ExtractDate = (Get-Date).ToString('yyyy-MM-dd')

## logentries config ##
$epoch   = Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
$AcctKey = "$($Kenobi.Boba)" #your logentries token
$LgEntry = "<logentries log set>"
$start   = [math]::truncate((Get-Date).AddHours(-1).ToUniversalTime().Subtract($epoch).TotalMilliSeconds)
$end     = [math]::truncate((Get-Date).ToUniversalTime().Subtract($epoch).TotalMilliSeconds)
## end logentries config ##

if (!(Test-Path "$workingDir\emps_reset.txt")) {Write-Verbose -Message "Nothing to move"} else{
    Move-Item "$workingDir\emps_reset.txt" -Destination "$archiveDir\emps_reset_$($ExtractDate)_$([System.IO.Path]::GetRandomFileName()).txt" -Force
}


irm -Uri "https:<logentries API>/$($AcctKey)/$($LgEntry)/?start=$($start)&end=$($end)" -Method GET -OutFile "$workingDir\emps_reset.txt"

if ( (Get-Content "$workingDir\emps_reset.txt").Length -gt 0 ) {
    $pattern= ‘\b[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}\b’
    $headers = 0..9
    $headers[7]="SlackUserID"
    $headers[8]="Slack_Username"
    $headers[9]="Text"
    Import-Csv "$workingDir\emps_reset.txt" -Header $headers -Delimiter "&" | `
    Select-Object -Property @{n='Slack_Username';e={$_.Slack_Username.Replace("user_name=","")}},@{n='Text';e={((($_.Text.Replace("text=","")).Replace("%3A",":")).Replace("%40","@"))}},@{n='SlackUserID';e={($_.SlackUserID).Replace("user_id=","")}} | `
    Where-Object {$_.Text -like '*mailto*'} | ForEach-Object -process { $sUser=((curl -F token="$($Slack)" -F user="$($_.SlackUserID)" https://slack.com/api/users.info -k --silent)|convertfrom-json).user.profile.email; [Regex]::Matches($_.Text, $pattern) | select Value,@{n='SlackUser';e={$sUser}}} |`
    Select-Object @{n='tableau_username';e={($_.Value).Replace("%7C","")}}, @{n='SlackUser';e={$_.SlackUser}}, @{n='tableau_slackuser';e={ $_.Value }} | Select-Object -Property @{n='tableau_username';e={($_.tableau_username).ToLower()}}, SlackUser,tableau_slackuser,@{n='Slack_Eq_PS';e={ $_.SlackUser -eq $_.tableau_slackuser }} | `
    #Where-Object {$_.Slack_Eq_PS -eq $true} | `
    Select-Object -ExpandProperty tableau_username -Unique | Out-File $psAddDir -Force ; & "AddEmps_With_Content.ps1"
} else {
        Write-Verbose -Message "Nothing to do. Exiting" -Verbose
    }