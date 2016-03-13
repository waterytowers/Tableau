
tabcmd logout
tabcmd login -s <your tableau server> -u <your ts username>  --password-file <location to pw file> --no-certcheck
tabcmd refreshextracts --datasource <data source> --project <project if not default> --no-certcheck 
tabcmd logout
