param(
    [Parameter(Mandatory = $true)][string]$DocumentPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [Parameter(Mandatory = $true)][string]$MarkerPrefix
)

# Inspect a private temporary document. Never attach to the user's open file.
$ErrorActionPreference = 'Stop'
$reportApp = $null
$reportDoc = $null
try {
    $failures = @()
    foreach ($programId in @('Word.Application', 'kwps.Application')) {
        try {
            $reportApp = New-Object -ComObject $programId
            $reportApp.Visible = $false
            $reportApp.DisplayAlerts = 0
            $reportDoc = $reportApp.Documents.Open($DocumentPath, $false, $true)
            $reportDoc.Repaginate()
            break
        } catch {
            $failures += ($programId + ': ' + $_.Exception.Message)
            try { if ($null -ne $reportDoc) { $reportDoc.Close(0) } } catch {}
            try {
                if ($null -ne $reportApp -and $reportApp.Documents.Count -eq 0) { $reportApp.Quit() }
            } catch {}
            $reportDoc = $null
            $reportApp = $null
        }
    }
    if ($null -eq $reportApp) {
        throw ('Could not paginate with Microsoft Word or WPS Writer. ' + ($failures -join ' | '))
    }
    $pageCount = [int]$reportDoc.ComputeStatistics(2)
    $positions = @{}
    $pageSetup = @{}
    foreach ($mark in $reportDoc.Bookmarks) {
        if ($mark.Name.StartsWith($MarkerPrefix, [System.StringComparison]::Ordinal)) {
            # 3 = wdActiveEndPageNumber: physical page, ignoring page-number restarts.
            $physicalPage = [int]$mark.Range.Information(3)
            $positions[$mark.Name] = $physicalPage
            $pageKey = [string]$physicalPage
            if (-not $pageSetup.ContainsKey($pageKey)) {
                $setup = $mark.Range.Sections.Item(1).PageSetup
                # Office lengths are points; DrawingML positions use EMU.
                $pageSetup[$pageKey] = @{
                    width = [long][Math]::Round([double]$setup.PageWidth * 12700)
                    height = [long][Math]::Round([double]$setup.PageHeight * 12700)
                    left = [long][Math]::Round([double]$setup.LeftMargin * 12700)
                    right = [long][Math]::Round([double]$setup.RightMargin * 12700)
                    top = [long][Math]::Round([double]$setup.TopMargin * 12700)
                    bottom = [long][Math]::Round([double]$setup.BottomMargin * 12700)
                    gutter = [long][Math]::Round([double]$setup.Gutter * 12700)
                    gutter_pos = [int]$setup.GutterPos
                    mirror = [bool]$setup.MirrorMargins
                }
            }
        }
    }
    if ($pageCount -lt 1 -or $positions.Count -eq 0) {
        throw 'Could not read document pagination and paragraph markers.'
    }
    $checks = @{}
    foreach ($shape in $reportDoc.Shapes) {
        if ($shape.Name.StartsWith('JRTrainingCheck_Page_', [System.StringComparison]::Ordinal)) {
            $checks[$shape.Name] = [int]$shape.Anchor.Information(3)
        }
    }
    @{ pages = $pageCount; positions = $positions; checks = $checks; page_setup = $pageSetup } |
        ConvertTo-Json -Depth 4 |
        Set-Content -LiteralPath $ResultPath -Encoding UTF8
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    try { if ($null -ne $reportDoc) { $reportDoc.Close(0) } } catch {}
    # WPS may share an application instance. Leave any other documents open.
    try { if ($null -ne $reportApp -and $reportApp.Documents.Count -eq 0) { $reportApp.Quit() } } catch {}
}
