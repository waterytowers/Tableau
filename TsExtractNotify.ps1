Set-StrictMode -Version Latest

$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList "<your username>", (Get-Content <path to pw file> | ConvertTo-SecureString)

Set-Location 'C:\TsAdmin\PsOutput'
Import-Module '<path to TsGetExtractModule>'

$connectionString = 'Driver={PostgreSQL ANSI(x64)};Server=<your server>; Port=8060; Database=workgroup; Uid=<your username>; Pwd=<your pw>;'
$query = @"
        select distinct
        started_at_pst,
        completed_at_pst, 
        finish_code,
        notes,
        title,
        (case when d.owner_name IS NULL then w.owner_name else d.owner_name end) as owner_name       
        from
	        (
	        SELECT
	        id, 
	        (started_at - interval '7 hours')   as started_at_pst, 
	        (completed_at - interval '7 hours') as completed_at_pst, 
	        finish_code, 
	        job_type, 
	        notes, 
	        job_name, 
	        title, 
	        subtitle,
	        rank() over(partition by title, job_name, started_at order by started_at desc) as rnk
	        FROM _background_tasks 
	        where 
			(started_at - interval '7 hours') >= ( (now()-interval '7 hours') - interval '3 hours' )
		    and 	completed_at IS NOT NULL
	        and 	subtitle IN('Workbook', 'Data Source') 
	        ) x     
	        left outer join _datasources d  on d.name = x.title 
		    left outer join _workbooks   w  on w.name = x.title 
        where 
            x.rnk      = 1
	  and  finish_code = 1
     

"@

Get-TsExtract -connectionString $connectionString -query $query | `
Select-Object -Skip 1 -Property started_at_pst, completed_at_pst, finish_code, title, extractfinish, owner_name, notes | `
Export-Csv -Path '<path to csv export>'    -NoTypeInformation -Delimiter ";" 

ForEach ( $extract in @(Import-Csv -Path <path to csv export> -Delimiter ";" | Select-Object -Property title, finish_code, owner_name, notes) ) {

    Try
     {
        #you're mailing the owner
        Send-MailMessage -To $extract.owner_name -From "<admin email>" `
         -Subject "$($extract.title) has Failed to refresh on Tableau Server" `
         -Body $extract.notes `
         -SmtpServer "<your email server>" `
         -Credential $credentials -UseSsl
     }
    Catch 
     {
        $_.Exception | select * | Out-File c:\TabCmd\TabCmdGetWbErr.txt
     }
}

##summary info sent to admin or group
Send-MailMessage -To "<admin email>" `
-From "<admin email>" `
-Subject "Extract Status Email(s) Sent" `
-Body "Total Extract Failure emails sent: $((Import-Csv C:\TsAdmin\PsOutput\TsExtracts.csv  -Delimiter ";" | Select title, owner_name | group -NoElement | select -expand count))" `
-SmtpServer "<your email server>" -Credential $credentials -UseSsl 

Set-Location $HOME
