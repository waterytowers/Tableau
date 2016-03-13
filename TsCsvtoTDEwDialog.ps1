#Requires -Version 3
Set-StrictMode -Version Latest

######################################################
# Developed by: Mike Roberts @ Pluralsight Data Team #
# Contact: mike-roberts@pluralsight.com              # 
# 8/1/15: Added support for true/false data chooser  #
######################################################


Add-Type -AssemblyName System.IO.Compression.Filesystem -ErrorAction Stop



##############################
## Edit with your variables ##
##############################

$TdeMaster      = gc '<some directory>\CsvTdeCreator_Master.txt' # this is the master TDS structure/config for use in the script. 
$TdeMasterTde   = '<some directory>\TdeMaster.tde' # this is the base/master/config TDE we'll use to package up with our datasource

$workingDirectory = '<some directory>\CsvtoTde'
$csvToTde         = '<your csv file>'
$separator        = '<your delimiter>'

$TsDatasource   = '<what you name it on Tableau Server>'
$ChooseDataType = $false # false will let Tableau do all the heavy lifting. Set to '$true' if you want to choose your data types. 
$TsServer       = '<your Tableau server>'
$TsUser         = '<Tableau admin user account>'
$TsPassword     = (<password file>)

$ExtractFolder  = '<some directory>\Masters\Data\Extracts'



######################################
## Nothing to change past this point #
######################################
icacls $workingDirectory /grant "<Tableau Server runas account>:(OI)(CI)(F)" # this is the Tableau Server runas account; yours will be different (it should) 

if (Test-Path "$workingDirectory" ) {
    Write-Verbose -Message "Directory Exists" ; Set-Location $workingDirectory
    } else {
        mkdir "$workingDirectory" ;Set-Location $workingDirectory
        }



if ($ChooseDataType) {
        if (Test-Path $csvToTde) {


                $myHeaders =  ipcsv $csvToTde -Delimiter $separator | select -first 1
                $myHeaders | Format-List | Out-String > $csvToTde.Replace(".csv",".txt")

                ######################################################################################### 
                # form script found here: https://technet.microsoft.com/en-us/%5Clibrary/Dn792464.aspx  #
                # Taking input for the columns and data types                                           #
                #########################################################################################

                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                $TsDataTypes = @()

                foreach($i in @( ipcsv $csvToTde.Replace(".csv",".txt") -Delimiter ":" -Header FieldNames | where {$_.FieldNames -notlike ''} )) {
                        $form = New-Object System.Windows.Forms.Form 
                        $form.Text = "Choose Data Type"
                        $form.Size = New-Object System.Drawing.Size(300,200) 
                        $form.StartPosition = "CenterScreen"

                        $OKButton = New-Object System.Windows.Forms.Button
                        $OKButton.Location = New-Object System.Drawing.Point(75,120)
                        $OKButton.Size = New-Object System.Drawing.Size(75,23)
                        $OKButton.Text = "OK"
                        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $form.AcceptButton = $OKButton
                        $form.Controls.Add($OKButton)

                        $CancelButton = New-Object System.Windows.Forms.Button
                        $CancelButton.Location = New-Object System.Drawing.Point(150,120)
                        $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                        $CancelButton.Text = "Cancel"
                        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                        $form.CancelButton = $CancelButton
                        $form.Controls.Add($CancelButton)

                        $label = New-Object System.Windows.Forms.Label
                        $label.Location = New-Object System.Drawing.Point(10,20) 
                        $label.Size = New-Object System.Drawing.Size(280,20) 
                        $label.Text = "Data type for $($i.FieldNames):"
                        $form.Controls.Add($label) 

                        $listBox = New-Object System.Windows.Forms.ListBox 
                        $listBox.Location = New-Object System.Drawing.Point(10,40) 
                        $listBox.Size = New-Object System.Drawing.Size(260,20) 
                        $listBox.Height = 80


                        [void] $listBox.Items.Add("string")
                        [void] $listBox.Items.Add("integer")
                        [void] $listBox.Items.Add("real")
                        [void] $listBox.Items.Add("date")
                        [void] $listBox.Items.Add("I don't know")


                        $form.Controls.Add($listBox) 

                        $form.Topmost = $True

                        $result = $form.ShowDialog()

                        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
                        {
                            $x = $listBox.SelectedItem
                            if ($x -like "I don't know") {$x="string"} else {$x}
                            $TsDataTypes += $($i.FieldNames).Trim()+";"+$x
                        }


                }

                $TsDataTypes > $csvToTde.Replace(".csv","_TableauFinal.csv")

                ### End form input section 


                if (Test-Path $workingDirectory) {
                    Set-Location $workingDirectory
                    } else {
                        mkdir $workingDirectory  ; Set-Location $workingDirectory
                       }

                #$myCsv = ipcsv $csvToTde -Delimiter $separator | select -First 1
                $directory = pwd | select -ExpandProperty path
                $relationName = $csvToTde.Replace(".","#")
                $relationTable = '['+$($csvToTde).Replace(".","#")+']'
                $finalTDE = $csvToTde.Replace(".csv","")


                $NewXml = @()
                $ord=0
                foreach($column in @( ipcsv $csvToTde.Replace(".csv","_TableauFinal.csv") -Header Field,Type -Delimiter ";" )) {
                    $ord++
                    $NewXml += "<column datatype='$($column.Type)' name='$($column.Field)' ordinal='$(($ord)-1)' />`n" 
                    }

                #columns swap
                $NewColumns = $TdeMaster | ForEach-Object { $_.replace("x_Replace_X","$($NewXml)") }
                $NewColumns | set-content "$workingDirectory\tdeTest0.tds" -Verbose 
                #directory swap
                $NewDirectory = (gc "$workingDirectory\tdeTest0.tds") | ForEach-Object {$_.replace("`'x_directory_X`'","'$($directory)'")} 
                $NewDirectory | set-content "$workingDirectory\tdeTest1.tds" -Verbose 
                #filename swap
                $NewFilename = (gc "$workingDirectory\tdeTest1.tds")  | ForEach-Object {$_.replace("`'x_filename_X`'","'$($csvToTde)'")} 
                $NewFilename | set-content "$workingDirectory\tdeTest2.tds" -Verbose
                #relationName swap
                $NewRelationName = (gc "$workingDirectory\tdeTest2.tds")  | ForEach-Object {$_.replace("`'x_relationName_X`'","'$($relationName)'")} 
                $NewRelationName | set-content "$workingDirectory\tdeTest3.tds" -Verbose
                #relationTable swap
                $NewRelationTable = (gc "$workingDirectory\tdeTest3.tds")  | ForEach-Object {$_.replace("`'[x_relationTable_X]`'","'$($relationTable)'")} 
                $NewRelationTable | set-content "$workingDirectory\tdeTest4.tds" -Verbose
                #extract swap
                $NewExtractLocation = (gc "$workingDirectory\tdeTest4.tds") |  ForEach-Object {$_.replace("`'x_Extract_X`'","'$($TdeMasterTde)'")} 
                $NewExtractLocation | set-content "$workingDirectory\tdeTest5.tds"
                #separator swap
                $NewSeperator = (gc "$workingDirectory\tdeTest5.tds") | ForEach-Object {$_.replace("`'x_seperator_X`'","'$($separator)'")}
                $NewSeperator | set-content "$workingDirectory\$($finalTDE).tds" -Verbose

                Remove-Item "$workingDirectory\tdeTest*.tds" -Verbose


                ####################################
                ## Now Publish the Tds to Tableau ##
                ####################################


                if (Test-Path $ExtractFolder ) {
                    Write-Verbose -Message "Directory Exists"
                    } else {
                        mkdir $ExtractFolder
                       }

                copy $TdeMasterTde -Destination $ExtractFolder -Verbose -Force

                if (Test-Path "$workingDirectory\$TsDatasource" ) {
                    Write-Verbose -Message "Directory Exists"
                    } else {
                        mkdir "$workingDirectory\$TsDatasource"
                       }

                # Zip and rename to tdsx
                copy "$workingDirectory\Masters\Data" -Recurse -Destination "$workingDirectory\$TsDatasource" -Verbose 
                copy "$workingDirectory\$($finalTDE).tds" -Destination "$workingDirectory\$TsDatasource" -Verbose

                [System.IO.Compression.ZipFile]::CreateFromDirectory("$workingDirectory\$TsDatasource","$workingDirectory\$TsDatasource.zip")
                if( (Test-Path "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx")) ) {
                        Remove-Item "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx") -Verbose
                    } else {Out-Null}
                Rename-Item "$workingDirectory\$TsDatasource.zip" -NewName "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx")
                Remove-Item "$workingDirectory\$TsDatasource" -Recurse -Verbose ; Remove-Item $ExtractFolder -Recurse -Verbose

                $TdeToServer =  "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx")
                $TdeNameOnServer = "$($csvToTde)".Replace(".csv","")


                if (Test-Path $TdeToServer) {
                    tabcmd login -s $TsServer -u $TsUser -p $TsPassword
                    tabcmd publish $TdeToServer -n "$($TdeNameOnServer)"  --project 'Default' --overwrite
                    tabcmd refreshextracts --datasource "$($TdeNameOnServer)" --project 'Default'
                    tabcmd logout
                } else {Write-Verbose -Message "Nothing to publish; exiting" -Verbose}

        } 
            else {Write-Verbose -Message "No csv datasource. Please enter valid file name" -Verbose} 
    }
 else {
        if (Test-Path $csvToTde) {
                if (Test-Path $workingDirectory) {
                    Set-Location $workingDirectory
                    } else {
                        mkdir $workingDirectory  ; Set-Location $workingDirectory
                       }

                #$myCsv = ipcsv $csvToTde -Delimiter $separator | select -First 1
                $directory = pwd | select -ExpandProperty path
                $relationName = $csvToTde.Replace(".","#")
                $relationTable = '['+$($csvToTde).Replace(".","#")+']'
                $finalTDE = $csvToTde.Replace(".csv","")


                $NewXml = " "
                #columns swap
                $NewColumns = $TdeMaster | ForEach-Object { $_.replace("x_Replace_X","$($NewXml)") }
                $NewColumns | set-content "$workingDirectory\tdeTest0.tds" -Verbose 
                #directory swap
                $NewDirectory = (gc "$workingDirectory\tdeTest0.tds") | ForEach-Object {$_.replace("`'x_directory_X`'","'$($directory)'")} 
                $NewDirectory | set-content "$workingDirectory\tdeTest1.tds" -Verbose 
                #filename swap
                $NewFilename = (gc "$workingDirectory\tdeTest1.tds")  | ForEach-Object {$_.replace("`'x_filename_X`'","'$($csvToTde)'")} 
                $NewFilename | set-content "$workingDirectory\tdeTest2.tds" -Verbose
                #relationName swap
                $NewRelationName = (gc "$workingDirectory\tdeTest2.tds")  | ForEach-Object {$_.replace("`'x_relationName_X`'","'$($relationName)'")} 
                $NewRelationName | set-content "$workingDirectory\tdeTest3.tds" -Verbose
                #relationTable swap
                $NewRelationTable = (gc "$workingDirectory\tdeTest3.tds")  | ForEach-Object {$_.replace("`'[x_relationTable_X]`'","'$($relationTable)'")} 
                $NewRelationTable | set-content "$workingDirectory\tdeTest4.tds" -Verbose
                #extract swap
                $NewExtractLocation = (gc "$workingDirectory\tdeTest4.tds") |  ForEach-Object {$_.replace("`'x_Extract_X`'","'$($TdeMasterTde)'")} 
                $NewExtractLocation | set-content "$workingDirectory\tdeTest5.tds"
                #separator swap
                $NewSeperator = (gc "$workingDirectory\tdeTest5.tds") | ForEach-Object {$_.replace("`'x_seperator_X`'","'$($separator)'")}
                $NewSeperator | set-content "$workingDirectory\$($finalTDE).tds" -Verbose

                Remove-Item "$workingDirectory\tdeTest*.tds" -Verbose


                ####################################
                ## Now Publish the Tds to Tableau ##
                ####################################


                if (Test-Path $ExtractFolder ) {
                    Write-Verbose -Message "Directory Exists"
                    } else {
                        mkdir $ExtractFolder
                       }

                copy $TdeMasterTde -Destination $ExtractFolder -Verbose -Force

                if (Test-Path "$workingDirectory\$TsDatasource" ) {
                    Write-Verbose -Message "Directory Exists"
                    } else {
                        mkdir "$workingDirectory\$TsDatasource"
                       }

                # Zip and rename to tdsx
                copy "$workingDirectory\Masters\Data" -Recurse -Destination "$workingDirectory\$TsDatasource" -Verbose 
                copy "$workingDirectory\$($finalTDE).tds" -Destination "$workingDirectory\$TsDatasource" -Verbose

                [System.IO.Compression.ZipFile]::CreateFromDirectory("$workingDirectory\$TsDatasource","$workingDirectory\$TsDatasource.zip")
                if( (Test-Path "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx")) ) {
                        Remove-Item "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx") -Verbose
                    } else {Out-Null}
                Rename-Item "$workingDirectory\$TsDatasource.zip" -NewName "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx")
                Remove-Item "$workingDirectory\$TsDatasource" -Recurse -Verbose ; Remove-Item $ExtractFolder -Recurse -Verbose

                $TdeToServer =  "$workingDirectory\$($csvToTde)".Replace(".csv",".tdsx")
                $TdeNameOnServer = "$($csvToTde)".Replace(".csv","")


                if (Test-Path $TdeToServer) {
                    tabcmd login -s $TsServer -u $TsUser -p $TsPassword
                    tabcmd publish $TdeToServer -n "$($TdeNameOnServer)"  --project 'Default' --overwrite
                    tabcmd refreshextracts --datasource "$($TdeNameOnServer)" --project 'Default'
                    tabcmd logout
                } else {Write-Verbose -Message "Nothing to publish; exiting" -Verbose}
        
       } else {Write-Verbose -Message "No csv datasource. Please enter valid file name" -Verbose} 
   } 
