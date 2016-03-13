Set-StrictMode -Version Latest

$hash = $null
$hash = [ordered]@{}
$admin = irm -Uri 'https://<yourTableauServer>/admin/systeminfo' -Method Get -UseDefaultCredentials | `
select-xml "//machines/*/*" | select -Expand node

foreach ($c in $admin){
 $hash.add($c.worker,$c.status)
    }

#Email portion
$email = $hash | Out-String

if($hash.GetEnumerator() | ? value -Like 'down') {
    Send-MailMessage -To '<email1>','<emailn>'  -Subject 'Server Status: Process Down Alert' -From '<YourEmail>' -Bcc '<bcc1>','<bccn>' `
    -Body $email -SmtpServer '<yourSmtpServer>'
}
else {
    Write-Verbose -Message "Server Status OK"
}
