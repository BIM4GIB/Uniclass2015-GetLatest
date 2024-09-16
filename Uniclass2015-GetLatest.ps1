#########################################################################################
#   ________  _______   ___      ___ ___  _________  _______   _______   ________       #  
#  |\   __  \|\  ___ \ |\  \    /  /|\  \|\___   ___\\  ___ \ |\  ___ \ |\   __  \      #  
#  \ \  \|\  \ \   __/|\ \  \  /  / | \  \|___ \  \_\ \   __/|\ \   __/|\ \  \|\  \     #  
#   \ \   _  _\ \  \_|/_\ \  \/  / / \ \  \   \ \  \ \ \  \_|/_\ \  \_|/_\ \   _  _\    #  
#    \ \  \\  \\ \  \_|\ \ \    / /   \ \  \   \ \  \ \ \  \_|\ \ \  \_|\ \ \  \\  \|   #  
#     \ \__\\ _\\ \_______\ \__/ /     \ \__\   \ \__\ \ \_______\ \_______\ \__\\ _\   #  
#      \|__|\|__|\|_______|\|__|/       \|__|    \|__|  \|_______|\|_______|\|__|\|__|  # 
#########################################################################################

$startDTM = (Get-Date)
$baseUrl = "https://uniclass.thenbs.com/download/downloadbundle?publishDate="
$currentDate = Get-Date
$publishMonths = @(1, 4, 7, 10)
$currentYear = $currentDate.Year
$possiblePublishDates = @()
$yearsToCheck = 1

for ($yearOffset = 0; $yearOffset -ge -$yearsToCheck; $yearOffset--) {
    $year = $currentYear + $yearOffset
    foreach ($month in $publishMonths) {
        $publishDate = Get-Date -Year $year -Month $month -Day 1
        if ($publishDate -le $currentDate) {
            $possiblePublishDates += $publishDate
        }
    }
}
$possiblePublishDates = $possiblePublishDates | Sort-Object -Descending
$latestPublishDate = $possiblePublishDates[0]
$publishDateString = $latestPublishDate.ToString("yyyy-MM")
$url = $baseUrl + $publishDateString
$desktopPath = Join-Path $Env:USERPROFILE "Desktop"
$destinationFolder = Join-Path $desktopPath "UniclassDownloads"
New-Item -ItemType Directory -Force -Path $destinationFolder | Out-Null
$destinationFile = Join-Path $destinationFolder "bundle_$($publishDateString).zip"

Write-Host "Attempting to download from $url"

try {
    Invoke-WebRequest -Uri $url -OutFile $destinationFile -ErrorAction Stop
    Write-Host "Downloaded file to $destinationFile"
    $extractDestination = Join-Path $destinationFolder "bundle_$($publishDateString)"
    Expand-Archive -Path $destinationFile -DestinationPath $extractDestination -Force
    Write-Host "Extracted contents to $extractDestination"
    $excelFiles = Get-ChildItem -Path $extractDestination -Filter "*.xlsx" -Recurse
    $excelFiles = $excelFiles | Where-Object {
        -not ($_.Name -like "*-codes-*" -or $_.Name -like "*-change-*")
    }
    if ($excelFiles.Count -eq 0) {
        Write-Warning "No Excel files found in $extractDestination after excluding specified files."
    } else {
        $codesToInclude = @("Ac", "Co", "EF", "En", "FI", "Ma", "PC", "PM", "Pr", "RK", "Ro", "SL", "Ss", "Te", "Zz")
        $mergedWorkbookPath = Join-Path $destinationFolder "merged_workbook.xlsx"
        if (Test-Path $mergedWorkbookPath) {
            Remove-Item $mergedWorkbookPath -Force
        }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $mergedWorkbook = $excel.Workbooks.Add()
        $sheetsCopied = $false
        $acSheetName = $null
        
        foreach ($file in $excelFiles) {
            Write-Host "Processing $($file.FullName)"
            $sourceWorkbook = $excel.Workbooks.Open($file.FullName)
            foreach ($sheet in $sourceWorkbook.Sheets) {
                $sheetName = $sheet.Name
                $includeSheet = $false
                foreach ($code in $codesToInclude) {
                    if ($sheetName -like "*$code*") {
                        $includeSheet = $true
                        break
                    }
                }
                if (-not $includeSheet) {
                    continue
                }
                $counter = 1
                $uniqueSheetName = $sheetName
                while ($true) {
                    $exists = $false
                    foreach ($ws in $mergedWorkbook.Sheets) {
                        if ($ws.Name -eq $uniqueSheetName) {
                            $exists = $true
                            break
                        }
                    }
                    if ($exists) {
                        $uniqueSheetName = "$($sheetName)_$counter"
                        $counter++
                    } else {
                        break
                    }
                }
                if (-not $sheetsCopied) {
                    $destinationSheet = $mergedWorkbook.Sheets.Item(1)
                    $sheet.UsedRange.Copy() | Out-Null
                    $destinationSheet.Paste() | Out-Null
                    $destinationSheet.Name = $uniqueSheetName
                    $sheetsCopied = $true
                } else {
                    $sheet.Copy([Type]::Missing, $mergedWorkbook.Sheets.Item($mergedWorkbook.Sheets.Count))
                    $mergedWorkbook.Sheets.Item($mergedWorkbook.Sheets.Count).Name = $uniqueSheetName
                }
                if ($uniqueSheetName -like "*Ac*") {
                    $acSheetName = $uniqueSheetName
                }
            }
            $sourceWorkbook.Close($false)
        }
        if ($mergedWorkbook.Sheets.Count -gt 1 -and $mergedWorkbook.Sheets.Item(1).Name -eq "Sheet1" -and -not $sheetsCopied) {
            $mergedWorkbook.Sheets.Item(1).Delete()
        }
        if ($acSheetName -ne $null) {
            $mergedWorkbook.Sheets.Item($acSheetName).Activate()
            Write-Host "Activated sheet '$acSheetName' before saving."
        } else {
            Write-Host "No 'Ac' sheet found to activate."
        }
        $mergedWorkbook.SaveAs($mergedWorkbookPath)
        Write-Host "Merged workbook saved to $mergedWorkbookPath"
        $mergedWorkbook.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
} catch {
    Write-Error "Failed to download from $url. Error: $_"
}
$endDTM = (Get-Date)
[System.Windows.Forms.MessageBox]::Show($this, "We are done here. Great Success!`n`nThat took about: $(($endDTM-$startDTM).totalseconds) seconds`n(give or take a millisecond)")
