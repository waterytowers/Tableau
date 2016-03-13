## this is the path , file that you're watching ##
$file = new-object System.IO.FileSystemWatcher `
      -ArgumentList @('C:\TsAdmin\PsOutput', 'TsDashboardResults.txt')

Register-ObjectEvent -inputobject $file -EventName Changed -SourceIdentifier FileChanged `
 -Action {## extract event test ##
            Set-StrictMode -Version Latest

            # login info as tabcmd doesn't take credential object
            $credentials = New-Object System.Management.Automation.PSCredential `
             -ArgumentList "<your Ts Username>", (Get-Content <path to secure string> | ConvertTo-SecureString)

            # get the workbooks from server
            Set-Location $HOME

            ## clearing old sessions
            tabcmd logout

            ## can also use password file for tabcmd
            tabcmd login -s <your server>`
             -u $credentials.UserName -p $credentials.GetNetworkCredential().Password  --no-certcheck 2>> c:\TabCmd\TabCmdLoginErr.txt
            tabcmd refreshextracts --datasource <data source to refresh> --project <project if not default> --no-certcheck 

            # cleanup and logout
            tabcmd logout 2>> C:\TabCmd\TabCmdLogoutErr.txt    
}
