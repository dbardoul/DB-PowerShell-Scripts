Clear-Host

Write-Host "Please select text or csv file with list of hostnames." -f Yellow

#Begin file explorer selection function
Function Get-OpenFile {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$filter,
        [string]$title,
        [bool]$dereferenceLinks,
        [bool]$multiSelect,
        [string]$grammar
    )
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
        Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = "c:\\"
    $OpenFileDialog.filter = $filter
    $OpenFileDialog.title = $title
    $OpenFileDialog.DereferenceLinks = $dereferenceLinks
    $OpenFileDialog.Multiselect = $multiSelect
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.ShowHelp = $true
    if ($grammar -eq "singular") { $OpenFileDialog.FileName }
    if ($grammar -eq "plural") { $OpenFileDialog.FileNames }
}
#End file explorer selection function

try {
    $InputFile = Get-OpenFile -filter "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv" `
                              -title "### Please select text or csv file with list of hostnames ###" `
                              -dereferenceLinks $true `
                              -multiSelect $false `
                              -grammar "singular"
    $computers = Get-Content $InputFile

    } catch {
        Write-Host "Text or csv file must be selected. Script ended." -f Red
        exit
}

Write-Host "File selected." -f Cyan

Write-Host "`nPlease select source file(s)." -f Yellow

$sourcefile = Get-OpenFile -filter "All files (*.*)| *.*" `
                           -title "### Please select source file(s) ###" `
                           -dereferenceLinks $false `
                           -multiSelect $true `
                           -grammar "plural"

if (!$sourcefile) {
    Write-Host "Source file(s) must be selected. Script ended." -f Red
    exit
}

Write-Host "Source file(s) selected." -f Cyan

Write-Host "`nSelect the destination location:" -f Yellow 
Write-Host "(1) " -f Yellow -NoNewline
Write-Host "C:\Users\Public\Desktop"
Write-Host "(2) " -f Yellow -NoNewline
Write-Host "C:\Temp"

do {
    $choice = Read-Host -Prompt "`nPlease choose an option"
    } while(($choice -lt 1) -or ($choice -gt 2))

switch ($choice) {
    1 { $destinationFolder = "C$\Users\Public\Desktop" }
    2 { $destinationFolder = "C$\temp" }
}

ForEach ($c in $computers) {
    $destination = "\\$c\" + $destinationFolder

    Write-Host ""

    ForEach ($s in $sourcefile) {

        try {

        #It will copy $sourcefile to the $destinationfolder. If the Folder does not exist it will create it.

        if (!(Test-Path -path $destination))
        {
            New-Item $destination -Type Directory
        }
        Copy-Item -Path $s -Destination $destination -Recurse | Out-Null

        Write-Host "Copied " -NoNewline
        Write-Host "$s " -f Cyan -NoNewline
        Write-Host "to " -NoNewline
        Write-Host "$c." -f Cyan

        } catch {

            Write-Host "Could " -NoNewline
            Write-Host "not " -f Red -NoNewline
            Write-Host "copy " -NoNewline
            Write-Host "$s " -f Cyan -NoNewline
            Write-Host "to " -NoNewline
            Write-Host "$c" -f Cyan

        }
    }
}