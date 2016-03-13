#Requires -Version 3
Set-StrictMode -Version Latest
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$ExtractDate = (Get-Date).ToString('yyyy-MM-dd')

############################################
# Update these variables
############################################

$export =  "<path where you want the csv files dropped to>" # these will be used to then add to database tables if you do that step
$tableauPath =  '<twb directory>'                           # location for all the TWB workbooks; ideally you've done the PostgreSQL part already. These can also be ad-hoc pulls

############################################
# End of changes
############################################



# Workbooks

foreach ($TableauTWB in @(Get-ChildItem -Path $tableauPath | Where-Object {$_.Extension -eq '.twb'} )) {
[xml]$wb =  (Get-Content  "$tableauPath\$TableauTWB") 

#  Need to add: 
#  actions
#  ability to pull out filter types via: "<tsfilter>+$($wb.workbook.worksheets.worksheet[0].table.view.filter.outerxml)+</tsfilter>"

    foreach ($i in $wb.workbook.worksheets.worksheet) {
        Clear-Variable sheetname -ErrorAction SilentlyContinue
        $sheetname = $i.name
        $i | Select-Object @{n='TwbName';e={$TableauTWB.BaseName}}, `
        @{n='Sheet';e={$sheetname}}, `
        @{n='SheetRows';e={$_.table.rows}}, `
        @{n='SheetSort';e={$_.table.view.sort.OuterXml}}, `
        @{n='SheetCols';e={$_.table.cols}}, `
        @{n='SheetFilters';e={$_.table.view.filter.column}}, `
        @{n='TotalSheetFilters';e={($_.table.view.filter.column).Count}}, `
        @{n='SheetFiltersType';e={($_.table.view.filter.class) | Select-Object -Unique}}, `
        @{n='SheetMarkType';e={$_.table.panes.pane.mark.class | Select-Object -Unique}}, `
        @{n='SheetEncodings';e={$_.table.panes.pane.encodings | Get-Member -MemberType Properties | select -expand name}}, `
        @{n='SheetEncodingsXml';e={$_.table.panes.pane.encodings.OuterXml}}, `
        @{n='SheetPageShelf';e={$_.table.pages.column}}, `
        #@{n='WorkbookShapes';e={$wb.workbook.external.shapes.shape | select @{n='ShapeName';e={$_.name}},@{n='ShapeBase64';e={$_.'#text'}} | ConvertTo-Json -Compress}}, `
        @{n='TotalTwbShapes';e={$wb.workbook.external.shapes.shape | group | select -expand count}}, `
        @{n='TotalTwbDataSources';e={$wb.workbook.datasources.datasource | Where-Object {$_.name -notlike 'parameters'} | group | select -expand count}}, `
        @{n='TotalTwbParameters';e={($wb.workbook.datasources.datasource | Where-Object {$_.name -like 'Parameters'}).column | group -Property name | Measure-Object -Property count -Sum | Select-Object -ExpandProperty Count}}, `
        @{n='TwbParametersType';e={($wb.workbook.datasources.datasource | Where-Object {$_.name -like 'parameters'}).column | Select-Object -Property datatype,'param-domain-type'}}, `
        @{n='ExtractDate';e={$ExtractDate}} | Export-Csv -Path "$export\TableauWBContent.csv" -Delimiter ";" -Append -Force -NoTypeInformation
     } 

}



# Dashboards


foreach ($TableauDashTWB in @(Get-ChildItem -Path $tableauPath | Where-Object {$_.Extension -eq '.twb'})) {
    [xml]$wbDash =  (Get-Content "$tableauPath\$TableauDashTWB") 

     foreach ($dash in $wbDash.workbook.dashboards.dashboard) {
        Clear-Variable myDash -ErrorAction SilentlyContinue ; Clear-Variable myDashSize -ErrorAction SilentlyContinue
        [xml]$myDash = $dash.outerxml; 
        $myDashSize = $dash.size | select max*,min* | group;
        $myDash | Select-XML -XPath "//dashboard//zones//zone" | select -expand node | where {$_.name -notlike 'zone'} | select @{n='TsWorkbook';e={$TableauDashTWB.BaseName}},@{n='dashname';e={$myDash.dashboard.name}},@{n='dashsize';e={$myDashSize.Group}},@{n='dashSheetName';e={$_.name}},param,type, `
        @{n='ExtractDate';e={$ExtractDate}} | Export-Csv -Path "$export\TableauDashContent.csv" -Delimiter ";" -Append -Force -NoTypeInformation
     }

 }




$Stopwatch.Stop()
"Total Script time: $($Stopwatch.Elapsed.TotalSeconds)" 