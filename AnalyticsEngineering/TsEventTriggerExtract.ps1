$computer = "$env:computername"
$filterNS = "root\cimv2"
$wmiNS = "root\subscription"

$query = @"
SELECT * FROM __InstanceOperationEvent WITHIN 30 WHERE TargetInstance ISA 'CIM_DataFile' 
AND TargetInstance.Drive = 'C:' AND TargetInstance.Path = '\\ETL\\'
"@

#ActiveScriptEventConsumer likes VBScript engine and using ScriptText b/c I don't want to lug around a vbs file
$scriptText = @"
Set objShell = CreateObject("Wscript.shell")
objShell.run("powershell -WindowStyle Hidden -executionpolicy bypass -file <your TabCmd refresh script goes here>")
"@


$filterPath = Set-WmiInstance -Class __EventFilter `
 -ComputerName $computer -Namespace $wmiNS -Arguments `
  @{name="NewRefresh"; EventNameSpace=$filterNS; QueryLanguage="WQL";
    Query=$query}


$consumerPath = Set-WmiInstance -Class ActiveScriptEventConsumer `
 -ComputerName $computer -Namespace $wmiNS `
 -Arguments @{name="ExtractTrigger"; ScriptText=$scriptText;
  ScriptingEngine="VBScript"}


#Binding
Set-WmiInstance -Class __FilterToConsumerBinding -ComputerName $computer `
  -Namespace $wmiNS -arguments @{Filter=$filterPath; Consumer=$consumerPath}
