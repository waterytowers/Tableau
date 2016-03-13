#Requires -Version 3
Set-StrictMode -Version Latest

$olFolderInbox = 6 # https://msdn.microsoft.com/en-us/library/office/aa210954(v=office.11).aspx
$outlook = new-object -comObject  outlook.application
$ns = $outlook.GetNameSpace("MAPI")
$inbox = $ns.GetDefaultFolder($olFolderInbox) 
$fileDownload = $inbox.Folders("Tableau")
#ns.SendAndReceive($true) #use this if you want to get data asap
$curdate = (Get-Date).ToString('yyyy-MM-dd')

# change these
$filepath = "<directory for csv's>"
$logPath = "<directory for logging email meta>"
# done with changes


if( (($fileDownload.Items | select attachments).attachments | where {$_.filename -like "*.csv"}).Count -ge 1) {
    
    foreach ($a in @(($fileDownload.Items | select attachments).attachments | where {$_.filename -like "*.csv"}))  {

            Write-Verbose -message  "Extracting file: $($a.FileName) to: $($filepath)"  -Verbose
            
            $a.SaveAsFile( "$($filepath)$($a.FileName)" )
            $a | Select-Object @{n='FileName';e={$a.FileName}},@{n='TsDatasource';e={($a.filename).Replace(".csv","")}},@{n='extractdate';e={$curdate}} | `
            export-csv -Path "$($logPath)\TsEmailMeta.csv" -Delimiter ";" -NoTypeInformation -Append

       } 
  } else {
            Write-Verbose -Message "No csv here" -Verbose
          }

$fileDownload.Items | Where-Object { (($_ | Select-Object attachments).attachments | Select-Object -expand filename) -like '*.csv' } | Select-Object Subject,SenderEmailAddress,SentOn,@{n='Attachments';e={ ($_ | select attachments).attachments | select -expand FileName }} | sort -property SentOn -desc

$outlook.Quit()
