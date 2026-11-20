param(
    [Parameter(Mandatory = $true)][string]$Template,
    [Parameter(Mandatory = $true)][string]$Output,
    [Parameter(Mandatory = $true)][string]$Payload
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function OleColor([int]$r, [int]$g, [int]$b) {
    return [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($r, $g, $b))
}

function Set-FormulaCell($cell, $value) {
    if ($null -eq $value) {
        $cell.Value2 = $null
    } elseif ($value -is [string] -and $value.StartsWith('=')) {
        $cell.Formula = $value
    } else {
        $cell.Value2 = $value
    }
}

function Write-Grid {
    param(
        [Parameter(Mandatory = $true)]$sheet,
        [Parameter(Mandatory = $true)][object[]]$headers,
        [Parameter(Mandatory = $true)][object[]]$rows,
        [string]$tableName,
        [bool]$makeTable
    )
    $headerArray = @($headers)
    $rowArray = @($rows)
    $rowCount = 1 + $rowArray.Count
    $columnCount = $headerArray.Count
    $grid = New-Object 'object[,]' $rowCount, $columnCount
    for ($c = 0; $c -lt $columnCount; $c++) {
        $grid[0, $c] = $headerArray[$c]
    }
    for ($r = 0; $r -lt $rowArray.Count; $r++) {
        $items = @($rowArray[$r])
        for ($c = 0; $c -lt $columnCount; $c++) {
            if ($c -lt $items.Count) { $grid[([int]$r + 1), $c] = $items[$c] }
        }
    }
    $range = $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($rowCount, $columnCount))
    $range.Value2 = $grid
    for ($r = 0; $r -lt $rowArray.Count; $r++) {
        $items = @($rowArray[$r])
        for ($c = 0; $c -lt $columnCount; $c++) {
            if ($c -lt $items.Count -and $items[$c] -is [string] -and $items[$c].StartsWith('=')) {
                $sheet.Cells(([int]$r + 2), ([int]$c + 1)).Formula = $items[$c]
            }
        }
    }
    $header = $sheet.Range($sheet.Cells(1, 1), $sheet.Cells(1, $columnCount))
    $header.Interior.Color = OleColor 47 84 150
    $header.Font.Color = OleColor 255 255 255
    $header.Font.Bold = $true
    $header.WrapText = $true
    $header.HorizontalAlignment = -4108
    $header.VerticalAlignment = -4108
    if ($makeTable -and $rowArray.Count -gt 0) {
        $list = $sheet.ListObjects.Add(1, $range, $null, 1)
        $list.Name = $tableName
        $list.TableStyle = 'TableStyleMedium2'
    }
    $sheet.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
    $sheet.UsedRange.Columns.ColumnWidth = 14
    return $range
}

function Add-GridSheet {
    param(
        [Parameter(Mandatory = $true)]$book,
        [Parameter(Mandatory = $true)][string]$name,
        [Parameter(Mandatory = $true)][object[]]$headers,
        [Parameter(Mandatory = $true)][object[]]$rows,
        [string]$tableName,
        [bool]$makeTable
    )
    $missing = [System.Type]::Missing
    $sheet = $book.Worksheets.Add($missing, $book.Worksheets.Item($book.Worksheets.Count))
    $sheet.Name = $name
    $sheet.Application.ActiveWindow.DisplayGridlines = $false
    Write-Grid -sheet $sheet -headers $headers -rows $rows -tableName $tableName -makeTable $makeTable | Out-Null
    return $sheet
}

$data = Get-Content -LiteralPath $Payload -Raw -Encoding UTF8 | ConvertFrom-Json
$stageLog = Join-Path (Split-Path -Parent $Output) 'po_usage_working_excel.stage.log'
function Stage([string]$message) { Add-Content -LiteralPath $stageLog -Value ("$(Get-Date -Format o) $message") }
Stage 'payload-loaded'
if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }
Copy-Item -LiteralPath $Template -Destination $Output -Force
Stage 'template-copied'

$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    Stage 'excel-created'
    $excel.Visible = $false
    Stage 'excel-hidden'
    $excel.DisplayAlerts = $false
    Stage 'alerts-set'
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    # External links are opened with UpdateLinks=0 below; keep the template's
    # cached values and request a full calculation when the user opens the copy.
    $book = $excel.Workbooks.Open($Output, 0, $false)
    Stage 'workbook-opened'

    foreach ($name in @('_Audit', '_PO_Master', '_Timesheet_Candidates', '_Monthly_Closing')) {
        foreach ($existing in @($book.Worksheets)) {
            if ($existing.Name -eq $name) { $existing.Delete() }
        }
    }
    Stage 'old-sheets-removed'

    $auditRows = @(
        @('Project', $data.project),
        @('Latest audit Run ID', $data.audit.run_id),
        @('Audit status', '=IF(OR(B9>0,B10>0,B11>0,B12>0,B13=0),"REVIEW_REQUIRED","READY_FOR_FREEZE")'),
        @('Hours per day', 8),
        @('', ''),
        @('PO PDF records', '=COUNTA(_PO_Master!B2:B10000)'),
        @('PO parsed records', '=COUNTIF(_PO_Master!C:C,"PARSED")'),
        @('PO parse review', '=COUNTIF(_PO_Master!C:C,"<>PARSED")'),
        @('Candidates missing PO', '=COUNTIF(_Timesheet_Candidates!G:G,"")'),
        @('Candidates requiring review', '=COUNTIF(_Timesheet_Candidates!M:M,"REVIEW")'),
        @('Monthly Closing unresolved mapping', '=COUNTIF(_Monthly_Closing!H:H,"UNRESOLVED_PROJECT")'),
        @('Approved timesheet facts', '=COUNTIF(_Timesheet_Candidates!K:K,"APPROVED")'),
        @('', ''),
        @('Gate meaning', 'Only READY_FOR_FREEZE may feed the official PO Usage Status; REVIEW_REQUIRED is a working copy.'),
        @('Source rule', 'PDF is the primary timesheet source after approval; Monthly Closing is reconciliation only and is never added again.')
    )
    $audit = Add-GridSheet -book $book -name '_Audit' -headers @('Metric', 'Value') -rows $auditRows -tableName '' -makeTable $false
    Stage 'audit-added'
    $audit.Range('A1:B1').Merge()
    $audit.Range('A1').Value2 = "$($data.project) - PO Usage Status Working Gate"
    $audit.Range('A1').Font.Bold = $true
    $audit.Range('A1').Font.Size = 14
    $audit.Range('A1').Font.Color = OleColor 47 84 150
    $audit.Columns.Item(1).ColumnWidth = 34
    $audit.Columns.Item(2).ColumnWidth = 105
    foreach ($row in @(2, 3, 4, 5, 16, 17)) { $audit.Cells($row, 1).Font.Bold = $true }
    $audit.Cells(4, 2).Interior.Color = OleColor 255 242 204

    $po = Add-GridSheet -book $book -name '_PO_Master' -headers @('Project', 'PO_ID', 'Parse_Status', 'PO_Date', 'SAP_PO_Number', 'Total_Excl_VAT_RMB', 'Delivery_Date', 'Source_File', 'Level_Quotas', 'Review_Flag') -rows $data.po_rows -tableName 'PO_Master_Table' -makeTable $true
    Stage 'po-added'
    if ($po.UsedRange.Rows.Count -gt 1) {
        $po.Range($po.Cells(2, 10), $po.Cells($po.UsedRange.Rows.Count, 10)).FormulaR1C1 = '=IF(RC[-7]="PARSED","OK","REVIEW")'
    }
    $po.Columns.Item(8).ColumnWidth = 44
    $po.Columns.Item(9).ColumnWidth = 80

    $ts = Add-GridSheet -book $book -name '_Timesheet_Candidates' -headers @('Candidate_ID', 'Raw_Name', 'Canonical_Name', 'Year', 'Month', 'Period', 'PO_ID', 'Hours', 'Days', 'Parse_Method', 'Review_Status', 'Source_File', 'Review_Flag') -rows $data.candidate_rows -tableName 'Timesheet_Candidates_Table' -makeTable $true
    Stage 'timesheet-added'
    if ($ts.UsedRange.Rows.Count -gt 1) {
        $last = $ts.UsedRange.Rows.Count
        $ts.Range($ts.Cells(2, 6), $ts.Cells($last, 6)).FormulaR1C1 = '=IF(AND(RC[-2]<>"",RC[-1]<>""),TEXT(RC[-2],"0000")&"-"&TEXT(RC[-1],"00"),"")'
        $ts.Range($ts.Cells(2, 9), $ts.Cells($last, 9)).FormulaR1C1 = '=IFERROR(RC[-1]/_Audit!R5C2,"")'
        $ts.Range($ts.Cells(2, 13), $ts.Cells($last, 13)).FormulaR1C1 = '=IF(OR(RC[-10]="",RC[-6]="",RC[-5]=""),"REVIEW",IF(RC[-2]="APPROVED","APPROVED","PENDING"))'
    }
    $ts.Columns.Item(12).ColumnWidth = 44

    $mc = Add-GridSheet -book $book -name '_Monthly_Closing' -headers @('Period', 'Sheet', 'Row', 'Raw_Name', 'Canonical_Name', 'Project_Number', 'Hours', 'Mapping_Status', 'Review_Flag') -rows $data.monthly_rows -tableName 'Monthly_Closing_Table' -makeTable $true
    Stage 'monthly-added'
    if ($mc.UsedRange.Rows.Count -gt 1) {
        $mc.Range($mc.Cells(2, 9), $mc.Cells($mc.UsedRange.Rows.Count, 9)).FormulaR1C1 = '=IF(RC[-1]="MAPPED","OK","REVIEW")'
    }
    $mc.Columns.Item(4).ColumnWidth = 26
    $mc.Columns.Item(5).ColumnWidth = 24

    $book.Worksheets.Item(1).Activate()
    # These calculation properties are not exposed as writable members in all
    # installed Excel versions.  The template was copied byte-for-byte, so its
    # existing formula caches remain intact; users can recalculate on open.
    Stage 'before-save'
    $book.Save()
    Stage 'saved'
    $book.Close($true)
    $excel.Quit()
    [Console]::WriteLine((ConvertTo-Json @{ ok = $true; output = $Output } -Compress))
}
finally {
    if ($book) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
