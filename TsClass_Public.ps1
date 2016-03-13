## enum has to be constant
Enum TsSites {
    yourSite1 
    yourSite2 
    yourSite3  
    yourSite4  
    yourSite5  

}
class tableau
{
    #[string]  $server
    [string]   $user
    [string[]] TsSite()        {return tabcmd listsites | ConvertFrom-String -Delimiter ":" | Where-Object {$_.P1 -like 'SITEID'} | Select-Object -Property P2}
    [string]   TabCmdVersion() {return tabcmd version}
    [TsSites]  $site
    [int]      TsDiskSpaceGB() {return Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object  @{n='FreeSpace';e={$_.FreeSpace/1GB}} | Select-Object -ExpandProperty FreeSpace}
    [string[]] TsCores()       {return Get-CimInstance -ClassName Win32_ComputerSystem | select NumberOfProcessors,NumberOfLogicalProcessors}
    [string]   TsOSBit()       {return Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty OSArchitecture}
    [string[]] TsRAM()         {return (Get-CimInstance -ClassName Win32_OperatingSystem  | Select-Object @{n='FreeRAM';e={$_.FreePhysicalMemory/1mb}},@{n='TotalRAM';e={$_.TotalVisibleMemorySize/1mb}} | Select-Object FreeRAM,TotalRAM) }
    
    [void] Login([string]$site)
    {        
        $Kenobi = (get-content "<location for your config file>") -join "`n" | ConvertFrom-StringData
        switch -CaseSensitive ($site)
        {
           "yourSite1" {$this.site=[TsSites]::yourSite1     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite2" {$this.site=[TsSites]::yourSite2     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite3" {$this.site=[TsSites]::yourSite3     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite4" {$this.site=[TsSites]::yourSite4     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite5" {$this.site=[TsSites]::yourSite5     ; tabcmd login -s $($Kenobi.VaderServer) -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ; break}
           "Default"   {$this.site=[TsSites]::Default       ; tabcmd login -s $($Kenobi.VaderServer) -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ; break}
        }

        
    }

      [void] Logout()
    {        
        tabcmd logout
    }

      [void] ExportTwbx([string]$workbook,[string]$filepath)
    {
        
        tabcmd get "/workbooks/$($workbook).twbx" -f "$filepath\$($workbook).twbx"
    }    

      [void] ExportTwbCsv([string]$workbook,[string]$view,[string]$filepath)
    {
        
        tabcmd get "/views/$($workbook)/$($view).csv" -f "$filepath\$($workbook).csv"
    }   


      [void] ExportTwbPng([string]$workbook,[string]$view,[string]$filepath)
    {
        
        tabcmd get "/views/$($workbook)/$($view)?:refresh=y" -f "$filepath\$($workbook).png"
    }  

     [void] MoveTwbxSite([string]$CurSite,[string]$workbook,[string]$NewSite)
    {
        $Kenobi = (get-content "<location for your config file>") -join "`n" | ConvertFrom-StringData
        switch -CaseSensitive ($CurSite)
        {
           "yourSite1" {$this.site=[TsSites]::yourSite1     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite2" {$this.site=[TsSites]::yourSite2     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite3" {$this.site=[TsSites]::yourSite3     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite4" {$this.site=[TsSites]::yourSite4     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite5" {$this.site=[TsSites]::yourSite5     ; tabcmd login -s $($Kenobi.VaderServer) -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ; break}
           "Default"   {$this.site=[TsSites]::Default       ; tabcmd login -s $($Kenobi.VaderServer) -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ; break}
        }
        
        tabcmd get "/workbooks/$($workbook).twbx" -f "$env:TEMP\$($workbook).twbx"
        
       
        switch -CaseSensitive ($NewSite)
        {
           "yourSite1" {$this.site=[TsSites]::yourSite1     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite2" {$this.site=[TsSites]::yourSite2     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite3" {$this.site=[TsSites]::yourSite3     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite4" {$this.site=[TsSites]::yourSite4     ; tabcmd login -s $($Kenobi.VaderServer) -t $this.site -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ;break}
           "yourSite5" {$this.site=[TsSites]::yourSite5     ; tabcmd login -s $($Kenobi.VaderServer) -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ; break}
           "Default"   {$this.site=[TsSites]::Default       ; tabcmd login -s $($Kenobi.VaderServer) -u $($Kenobi.VaderUsername) -p $($Kenobi.VaderPassword) ; break}
        }
        tabcmd publish "$env:TEMP\$($workbook).twbx" --overwrite



    }  
 
}

# Instantiate the class
$tableau = [tableau]::new()

# Some examples
$tableau.Login("Default")
$tableau.ExportTwbx('MycoolWorkbook','<some file path>')
$tableau.MoveTwbxSite("OldSite","MyOtherCoolWorkbook","NewSite")
$tableau.TsDiskSpace()
$tableau.TsCores()
$tableau.ExportTwbPng("MycoolWorkbook","MycoolWorkbookView","<some file path>")
$tableau.Logout()
