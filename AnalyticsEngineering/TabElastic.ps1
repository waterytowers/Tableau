#Requires -Version 3
Set-StrictMode -Version 3

$tabElasticStart = Get-Date
$curBackGr = gcim -ClassName win32_process -Filter "name='backgrounder.exe'" 
$Cores = gcim -ClassName win32_processor | select -expand NumberOfCores

$tabElastic = @{'maxBackGr' = $Cores/2;
               'minBackGr' = $Cores/4
                }

switch ($curBackGr.Count) 
         {
            4 {$TsBackGr=2;Break}
            2 {$TsBackGr=4;Break}
         } 


if ($curBackGr.count -ge $tabElastic.maxBackGr) {
       Write-Verbose -Message "Backgrounders set at $($curBackGr.count). Setting to $($TsBackGr)" -Verbose
$pinvokes = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr FindWindow(string className, string windowName);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
'@

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -MemberDefinition $pinvokes -Name NativeMethods -Namespace MyUtils



        tabadmin stop 
        tabadmin cleanup
        tabconfig
        Start-Sleep -Seconds 10

        # Win 1: Tableau Server Configuration
        $hwnd = [MyUtils.NativeMethods]::FindWindow("#32770","Tableau Server Configuration") # need class name and window title
        [MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
        [System.Windows.Forms.SendKeys]::SendWait("^+{TAB}")
        [System.Windows.Forms.SendKeys]::SendWait("{RIGHT 2}")
        [System.Windows.Forms.SendKeys]::SendWait("{LEFT 6}") 
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 3}") 
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}") 
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 2}") 
        [System.Windows.Forms.SendKeys]::SendWait("$($TsBackGr)")
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 7}")
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}") 
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 4}")
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER 2}")
        

        while( (gcim -ClassName win32_process -filter "name='tabconfig.exe'").ProcessId -gt 1) {
            Write-Verbose -Message "Waiting for tabconfig to close" -Verbose
            Start-Sleep -Seconds 1
            } 

        tabadmin start
        tabadmin status -v

        gcim -class win32_operatingsystem | select freephysicalmemory,totalvisiblememorysize,@{n='per';e={1-($_.freephysicalmemory/$_.totalvisiblememorysize)}}
        
        $tabElasticEnd = Get-Date
}
    
     else {

      Write-Verbose -Message "Backgrounders set at $($curBackGr.count). Setting to $($TsBackGr)" -Verbose
$pinvokes = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr FindWindow(string className, string windowName);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
'@

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -MemberDefinition $pinvokes -Name NativeMethods -Namespace MyUtils
        

        tabadmin stop 
        tabadmin cleanup
        tabconfig
        Start-Sleep -Seconds 10

        # Win 1: Tableau Server Configuration
        $hwnd = [MyUtils.NativeMethods]::FindWindow("#32770","Tableau Server Configuration") # need class name and window title
        [MyUtils.NativeMethods]::SetForegroundWindow($hwnd)
        [System.Windows.Forms.SendKeys]::SendWait("^+{TAB}")
        [System.Windows.Forms.SendKeys]::SendWait("{RIGHT 2}")
        [System.Windows.Forms.SendKeys]::SendWait("{LEFT 6}") 
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 3}") 
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}") 
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 2}") 
        [System.Windows.Forms.SendKeys]::SendWait("$($TsBackGr)")
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 7}")
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}") 
        [System.Windows.Forms.SendKeys]::SendWait("{TAB 4}")
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER 2}")

        
        while( (gcim -ClassName win32_process -filter "name='tabconfig.exe'").ProcessId -gt 1) {
            Write-Verbose -Message "Waiting for tabconfig to close" -Verbose
            Start-Sleep -Seconds 1
            } 

        tabadmin start
        tabadmin status -v

        gcim -class win32_operatingsystem | select freephysicalmemory,totalvisiblememorysize,@{n='per';e={1-($_.freephysicalmemory/$_.totalvisiblememorysize)}}
        $tabElasticEnd = Get-Date
}



$tabScriptTime = ($tabElasticEnd - $tabElasticStart).TotalSeconds
"Script took $($tabScriptTime) seconds."
