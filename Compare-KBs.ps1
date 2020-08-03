Clear-Host

Write-Host "Please select text or csv file with list of hostnames." -f Yellow

#Begin Function file explorer pop-up selection
Function Get-OpenFile { 
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$title
    )
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
        Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = "c:\\"
    $OpenFileDialog.filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv"
    $OpenFileDialog.title = $title
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
    $OpenFileDialog.ShowHelp = $true
    }
#End Funtion file explorer pop-up selection

try {    
    $InputFile = Get-OpenFile -title "### Please select text or csv file with list of hostnames ###"
    $computers = Get-Content $InputFile

    } catch {
        Write-Host "Text or csv file must be selected. Script ended." -f Red
        exit
}

Write-Host "File selected." -f Cyan

Write-Host "`nPlease select text or csv file with list of KBs." -f Yellow

try {
    $InputFile = Get-OpenFile -title "### Please select text or csv file with list of KBs ###"
    $listKBs = Get-Content $InputFile

    } catch {
        Write-Host "Text or csv file must be selected. Script ended." -f Red
        exit
}

Write-Host "File selected." -f Cyan

Write-Host "`nPlease select location to store report." -f Yellow

#Begin Function folder explorer pop-up selection
Function Get-OpenFolder {
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') |
        Out-Null

    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.RootFolder = 'MyComputer'
    $FolderBrowserDialog.Description = "### Please select location to store report ###"
    $FolderBrowserDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })) | Out-Null
    return $FolderBrowserDialog.SelectedPath
}
#End Funtion folder explorer pop-up selection

$OutputFile = Get-OpenFolder
if (!$OutputFile) {
    Write-Host "Location must be selected. Script ended." -f Red
    exit
}

Write-Host "Location selected." -f Cyan

Write-Host "`nCreating KB Report spreadsheet in $OutputFile... Please wait..." -ForegroundColor Yellow

<# Excel spreadsheet creation credit goes to Boe Prox from https://learn-powershell.net/
   Link - https://learn-powershell.net/2012/12/20/powershell-and-excel-adding-some-formatting-to-your-report/
#>

#Create excel COM object
$excel = New-Object -ComObject excel.application

#Make Visible
$excel.Visible = $False

#Add a workbook
$workbook = $excel.Workbooks.Add()

#Connect to first worksheet to rename and make active
$serverInfoSheet = $workbook.Worksheets.Item(1)
$serverInfoSheet.Name = 'KB Report'
$serverInfoSheet.Activate() | Out-Null

#Create a Title for the first worksheet and adjust the font
$row = 1
$Column = 1
$serverInfoSheet.Cells.Item($row,$column)= 'KB Report'

$serverInfoSheet.Cells.Item($row,$column).Font.Size = 18
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$serverInfoSheet.Cells.Item($row,$column).Font.Name = "Cambria"
$serverInfoSheet.Cells.Item($row,$column).Font.ThemeFont = 1
$serverInfoSheet.Cells.Item($row,$column).Font.ThemeColor = 4
$serverInfoSheet.Cells.Item($row,$column).Font.ColorIndex = 55
$serverInfoSheet.Cells.Item($row,$column).Font.Color = 8210719

$range = $serverInfoSheet.Range("a1","c2")
$range.Merge() | Out-Null
$range.VerticalAlignment = -4160

#Increment row for next set of data
$row++;$row++

$serverInfoSheet.Cells.Item($row,$column)= 'If cell is green, there was a match'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =4
$range = $serverInfoSheet.Range("a3","c3")
$range.Merge() | Out-Null

$row++

$serverInfoSheet.Cells.Item($row,$column)= 'If cell is red, there was not a match'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =3
$range = $serverInfoSheet.Range("a4","c4")
$range.Merge() | Out-Null

$row++

#Create a header for Hostname Report; set each cell to Bold and add a background color
$serverInfoSheet.Cells.Item($row,$column)= 'Hostname'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'KB'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True

#Increment Row and reset Column back to first column
$row++
$Column = 1

#Save the initial row so it can be used later to create a border
$initalRow = $row

#Loop through list of computers and gather data for each row
ForEach ($c in $computers) {
    
    $HostKBs = Get-HotFix -ComputerName $c | sort-object Hotfixid | ForEach-Object {$_.Hotfixid}

    #Hostname
    $serverInfoSheet.Cells.Item($row,$column) = $c.ToUpper()
    $Column++

    ForEach ($k in $listKBs) {

        $serverInfoSheet.Cells.Item($row,$column)= $k

        if ($HostKBs -eq $k) {
            $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =4
        } else {
            $serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =3
        }
        $row++
    }

    $dataRange = $serverInfoSheet.Range(("A{0}" -f $initalRow),("B{0}" -f $row))
    $dataRange.BorderAround(1) | Out-Null

    #Increment to next row and reset Column to 1
    $Column = 1
    $row++
}

#Auto fit everything so it looks better
$usedRange = $serverInfoSheet.UsedRange	
$usedRange.EntireColumn.AutoFit() | Out-Null
$usedRange.EntireRow.AutoFit() | Out-Null

#Timestamp variable for saved file name
$timestamp = $(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmss"))

#Save the file
$workbook.SaveAs($OutputFile + "\KB_Report_$timestamp.xlsx")

#Quit the application
$excel.Quit()

#Release COM Object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null

Write-Host "KB Report complete." -f Cyan