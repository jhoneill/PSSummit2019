cd $env:TEMP

#first examples use get-netadapter - not available in all versions of PowerShell, but you can use something elese 

#1. Simplest export - picks some defaults, picks a random file name, opens excel.
Get-NetAdapter | Export-Excel

#2. Specify the file name, and just export the data, don't open excel
Get-NetAdapter | Sort-Object  MacAddress |
    Select-Object   MacAddress, Status, LinkSpeed, MediaType |
        Export-excel -path Net1.xlsx

start .\Net1.xlsx


#3. Specify the file name, make it look a bit nicer  and open excel
Get-NetAdapter | Sort-Object  MacAddress |
    Select-Object   MacAddress, Status, LinkSpeed, MediaType |
        Export-excel -path Net2.xlsx -Show -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow


#4. This time don't treat 802.3 as a number
Get-NetAdapter | Sort-Object  MacAddress |
        Select-Object   MacAddress, Status, LinkSpeed, MediaType |
            Export-excel -path Net3.xlsx -Show -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -NoNumberConversion MediaType

#5. conditional formatting. When the file is opened you can click the filter option and filter by color.
$cf = New-ConditionalText -Text "up" -ConditionalTextColor Green -BackgroundColor White -PatternType None
Get-NetAdapter | Sort-Object  MacAddress |
        Select-Object   MacAddress, Status, LinkSpeed, MediaType |
            Export-excel -path Net4.xlsx -Show -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -ConditionalFormat $cf

#While we're with conditional formatting, let's do some more and see PASSTHRU

#We want to get some data for the next part - use get-sql and query the database of F1 results.
#You can get the database from  https://1drv.ms/f/s!AhfYu7-CJv4egbt5FD7Cdxi8jSz3aQ and installModule GetSQL 

Get-SQL -Excel -Connection C:\Users\mcp\onedrive\public\f1\f1Results.xlsx -Session f1

#Get-Sql -excel -connection "xlfile"  sets up a SQL connection to an excel file (-Access makes it an access file, -MSSQLServer makes it a sql box, and anything else is an ODBC string)
# -Session xxx allows multiple connections to be opened at once and defines xxx as an alias for Get-SQL <<session xxx>> 
# Main reason for using get-sql is we can build up the table as below. RacePos (position) is a string (includes DNF, retired etc.) 

#6. get wins per driver, filter out < 5 wins, use flag to show more/less successful
f1 -Table "Results" -GroupBy "DriverName" -Select "DriverName","Count (RaceDate) as wins" -Where "RacePos" -eq "'1'" -Verbose |
    Select-Object -Property drivername,wins | Where-Object -Property wins -ge 5 |
        Export-Excel -path f1winners.xlsx -AutoNameRange  -AutoSize
$excel = Open-ExcelPackage -Path .\f1winners.xlsx
Add-ConditionalFormatting  -Address $excel.Sheet1.Cells["Wins"] -ThreeIconsSet Flags
Close-ExcelPackage -Show -ExcelPackage $excel

#7. Didn't like that because very few are > 2/3 Max score in even 1/3 ... so lets take control of which values get which flags

del .\f1winners.xlsx
$excel = f1 -Table "Results" -GroupBy "DriverName" -Select "DriverName","Count (RaceDate) as wins" -Where "RacePos" -eq "'1'"  |
    Select-Object -Property drivername,wins |  Where-Object -Property wins -ge 5 | Sort-Object -Property wins -Descending |
        Export-Excel -path f1winners.xlsx -AutoNameRange  -AutoSize   -PassThru

$cf =  Add-ConditionalFormatting  -Address $excel.Sheet1.Cells["Wins"] -ThreeIconsSet Flags  -PassThru
#What does conditional formatting look like 
$cf
$cf.Icon2
#OK ... lets make those percentiles so we get even thirds
$cf.icon2.type="Percentile"
$cf.icon3.type="Percentile"

Close-ExcelPackage -Show -ExcelPackage $excel

#Leave things tidy
f1 -Close

#A different (and bigger) database - My pictures from Adobe Lightroom
####### For the download I've provided pictureinfo.csv. just do $pictureInfo = import-csv pictureinfo.csv 

$SQL =@"
SELECT      rootfile.baseName || '.' || rootfile.extension       AS fileName,
            metadata.dateDay         , metadata.dateMonth, metadata.dateYear,
            image.fileFormat         ,
            Image.captureTime       AS dateTaken ,
            metadata.hasGPS          ,
            metadata.focalLength     ,
            metadata.Aperture       AS apertureValue,
            metadata.ISOSpeedRating AS ISOSpeed,
            metadata.ShutterSpeed   AS shutterSpeedValue,
            Camera.Value            AS cameraModel,
            LensRef.value           AS lensModel
FROM        Adobe_images               image
JOIN        AgLibraryFile              rootFile ON   rootfile.id_local =  image.rootFile
JOIN        AgharvestedExifMetadata    metadata ON      image.id_local =  metadata.image
LEFT JOIN   AgInternedExifLens         LensRef  ON    LensRef.id_Local =  metadata.lensRef
LEFT JOIN   AgInternedExifCameraModel  Camera   ON     Camera.id_local =  metadata.cameraModelRef
ORDER BY   fileName
"@

#Just use the default session this time
Get-SQL -Connection "DSN=LR"
#PowerShell has an implied alias xxx for get-xxx so SQL = Get-SQL ; get-sql takes -SQL parameter and we've put the query in $sql so I run "sequel sequel sequel" 
$pictureInfo = SQL -SQL $sql
Get-SQL -Close

#I have "HowLong" in my profile. A bit like measure-command but can be called on ay previous command(s) 
<#
    Function HowLong {
    param   (
        #The history ID of the command. If not specified the last command to run is used  
        [Parameter(ValueFromPipeLine=$true)]$id =  ($MyInvocation.HistoryId -1) 
    )
    process { foreach ($i in $id) { (get-history -Id $i -count 1 ).endexecutiontime.subtract((get-history -id $i -count 1).startexecutiontime).totalseconds } }        
    }
#>

#8. 18 cols * 4800 rows ~ 86,000 cells  -note extra row stuff, and we parsed numbers but not dates
Export-Excel -InputObject $pictureInfo
HowLong



#9. Data rows are ugly - will fix that one day - for now use excludeProperty, fix the date - add a pivot table
$pictureinfo | Select-Object -Property *,@{n="taken";e={[datetime]$_.datetaken}} |
    Export-Excel -ExcludeProperty datetaken,RowError,RowState,Table,ItemArray,HasErrors -IncludePivotTable `
     -PivotRows cameraModel -PivotColumns lensmodel -PivotData @{'Taken'='Count'} -IncludePivotChart -ChartType ColumnStacked

#10. Unwieldy command line so use definitions - which have more options - we can also use Send-SqlDataToExcel which bolts Get-SQL logic onto Export-Excel 
$cDef = New-ExcelChartDefinition -title "Lens and Camera Usage" -ChartType ColumnStacked -Row 17 -Column 0 -Width 750 -height 500
$pdef = New-PivotTableDefinition -PivotTableName PicturePivot -PivotRows cameraModel -PivotColumns Lensmodel -PivotData  @{'dateTaken'='Count'}  -PivotChartDefinition $cDef -Activate

$cdef
$pdef
$pdef.PicturePivot

<# If you have downloaded the csv use -- BE PATIENT this way is a lot slower than send-sqldata and takes ~60 seconds on my machine ! 
$pdef = New-PivotTableDefinition -PivotTableName PicturePivot -PivotRows cameraModel -PivotColumns Lensmodel -PivotData  @{'Taken'='Count'}  -PivotChartDefinition $cDef -Activate
$excel = $pictureinfo | Select-Object -Property *,@{n="taken";e={[datetime]$_.datetaken}} -first 1000 |
  Export-Excel -ExcludeProperty datetaken,RowError,RowState,Table,ItemArray,HasErrors -path picturedemo.xlsx -PivotTableDefinition $pdef -AutoNameRange -PassThru
#>

$excel  = Send-SQLDataToExcel -Connection "DSN=LR" -SQL $SQL -path picturedemo.xlsx -PivotTableDefinition $pdef -AutoNameRange -PassThru

howlong

#11. We can access the sheet by name ... and access its cells or ranges of cells in multiple ways
$sheet = $excel.Sheet1
#how wide is the sheet
$sheet.Dimension
$sheet.cells["A1"]
$sheet.cells[1,2]

#12. find a column header and set that column - here we're going to set the number of decimals on av and tv
$col = 1
while ($col -lt $sheet.Dimension.Columns -and $sheet.cells[1,$col].Value -ne "apertureValue") {$col++}
$col
Set-ExcelColumn -worksheet $sheet -Column $col  -NumberFormat "0.00"
$sheet.cells[2,$col].Style.Numberformat

#13 better
1..$sheet.Dimension.Columns | where {$sheet.cells[1,$_].value -in @("apertureValue","shutterSpeedValue") } |
    Set-ExcelColumn -worksheet $sheet  -NumberFormat "0.00"

#14. add columns ... use auto created ranges in the formula (data had "apertureValue" and "shutterSpeedValue" columns, they are now ranges)
Set-ExcelColumn -Worksheet $sheet -Heading "f. stop"       -Value "=SQRT(POWER(2,apertureValue))" -NumberFormat '"f/"0.0'
Set-ExcelColumn -Worksheet $sheet -Heading "Exposure time" -Value "=1/(SQRT(POWER(2,shutterSpeedValue)))" -NumberFormat 'Fraction'

#how wide is it *now*
$sheet.Dimension
#this tries to render the cells, which is slow and WINDOWS ONLY
$sheet.Cells[$sheet.Dimension.Address].AutoFitColumns()

#15 wouldn't it look nicer as a table ?
Add-ExcelTable -Range $sheet.cells[$($sheet.Dimension.address)] -TableStyle Light1 -TableName "Pictures"

#16. Add save. Can do that with Close-ExcelPackage .. but didn't fix the top row do that now -it's an export option so use export, close and show the file. 
Export-Excel -ExcelPackage $excel -WorksheetName sheet1 -FreezeTopRow -Show


##### Drum roll - last minute code from doug -- full of pivotty goodness. 
## sa.ps1 file in the download - use your own paths 

#if I'm not in VSCode this is where I need to load script analyzer from 
#ipmo C:\Users\mcp\.vscode\extensions\ms-vscode.powershell-preview-2.0.2\modules\PSScriptAnalyzer\1.18.0\PSScriptAnalyzer.psd1

psedit C:\Users\mcp\Documents\GitHub\SA.ps1
cd 'C:\Program Files\WindowsPowerShell\Modules\ImportExcel\5.0.1\'
C:\Users\mcp\Documents\GitHub\SA.ps1 -xlfile "$env:temp\ImportExcel-5.0.1.xlsx"

cd C:\Users\mcp\Documents\GitHub\ImportExcel
C:\Users\mcp\Documents\GitHub\SA.ps1 -xlfile "$env:temp\GitHub-ImportExcel.xlsx"

Cd $env:TEMP

#OK back to our schedule  what about about simpler lists ? Lets use some lorem-ipsum data
#Lorem-ipsum is dull - lets use some from http://www.cupcakeipsum.com

#CupCake-ipsum.txt is in the download set, just change the path below

#17. Insert, then insert to the right , repeat
$cupcake = Get-Content "C:\Users\mcp\Documents\WindowsPowerShell\cupcake-ipsum.txt" -Encoding UTF8
$cupcake.count
$col = 1 ;
foreach ($c in $cupcake) {
        $c -split "\s+" | Export-Excel -path cupcake1.xlsx -StartColumn $col
        $col++
}

#Have a look and make subtle changes (save as cupcake2) - we'll come back to that.
 start cupcake1.xlsx

#18. BTW  Yes we can append.
$cupcake[0] -split "\s+" | Export-Excel -path cclong.xlsx
 foreach ($c in $cupcake) {$c -split "\s+" | export-excel -path cclong.xlsx -append   }

 start cclong.xlsx

#19. A word of warning about this data. It doesn't have a header
 import-excel .\cclong.xlsx | select -First 10

import-excel -NoHeader .\cclong.xlsx -EndRow 10

import-excel .\cupcake1.xlsx -EndRow 10 | ft

import-excel .\cupcake1.xlsx -NoHeader -EndRow 10 | ft

#20  lets compare the two wide sheets (we need to remember no header) Mark up the changed rows - try filtering to the green ones !

compare-worksheet -Referencefile .\cupcake1.xlsx -Differencefile .\cupcake2.xlsx -NoHeader -BackgroundColor LightGreen -FontColor Red -Show

#21 if we can compare sheets how about merging them ? first a simple one
merge-Worksheet -Referencefile .\GitHub-ImportExcel.xlsx -Differencefile .\ImportExcel-5.0.1.xlsx -WorksheetName breakdown -Startrow 2 -OutputFile versions.xlsx  -Show -key "Row labels"

#22 then a whole collection of them

dir net*.xlsx |  Merge-MultipleSheets -OutputFile netmetge.xlsx -WorkSheetName sheet1   -HideRowNumbers -OutputSheetName network  -Key macaddress  -ExcludeProperty mediatype -Show

#23  this last one is from my msftgraph module. Install that and you should have the example 
psedit C:\Users\mcp\Documents\WindowsPowerShell\Modules\MsftGraph\Examples\Export-planner-to-xlsx.ps1

