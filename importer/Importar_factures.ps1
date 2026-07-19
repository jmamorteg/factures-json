[CmdletBinding()]
param(
    [Parameter(Position = 0)] [string]$JsonPath,
    [Parameter(Position = 1)] [string]$WorkbookPath,
    [string]$WorksheetName = "Entrada_Factures",
    [string]$TableName = "tblEntradaFactures",
    [switch]$NoBackup,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$exitCode = 0
$excel = $null
$workbook = $null
$worksheet = $null
$table = $null

function Pick-File {
    param([string]$Title, [string]$Filter)
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.CheckFileExists = $true
    try {
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            throw "Operacio cancel-lada."
        }
        return $dialog.FileName
    }
    finally { $dialog.Dispose() }
}

function Release-Com {
    param($Object)
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch { }
    }
}

function Get-ColumnIndex {
    param($Table, [string[]]$Names, [switch]$Optional)
    foreach ($name in $Names) {
        $column = $null
        try {
            $column = $Table.ListColumns.Item($name)
            return [int]$column.Index
        }
        catch { }
        finally { Release-Com $column }
    }
    if ($Optional) { return $null }
    throw "Falta una columna obligatoria: $($Names -join ' / ')"
}

function To-Number {
    param($Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    if ($Value -is [ValueType]) { return [double]$Value }

    $text = ([string]$Value).Trim()
    $number = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
        return $number
    }
    throw "Valor numeric no valid: $Value"
}

function To-IsoDate {
    param($Value)
    if ($Value -is [DateTime]) { return $Value.ToString("yyyy-MM-dd") }
    if ($Value -is [ValueType]) { return [DateTime]::FromOADate([double]$Value).ToString("yyyy-MM-dd") }

    $text = ([string]$Value).Trim()
    $date = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($text, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
        return $date.ToString("yyyy-MM-dd")
    }
    if ([DateTime]::TryParse($text, [Globalization.CultureInfo]::CurrentCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
        return $date.ToString("yyyy-MM-dd")
    }
    throw "Data no valida: $Value"
}

function Normalize-Key {
    param($Value)
    if ($null -eq $Value) { return "" }
    return (([string]$Value).Trim().ToUpperInvariant() -replace '\s+', '')
}

function Invoice-Key {
    param($Tipus, $NomAlias, $NumFactura, $Data, $Bi)
    $amount = (To-Number $Bi).ToString("0.#####", [Globalization.CultureInfo]::InvariantCulture)
    return @(
        (Normalize-Key $Tipus),
        (Normalize-Key $NomAlias),
        (Normalize-Key $NumFactura),
        (To-IsoDate $Data),
        $amount
    ) -join "|"
}

function Set-Text {
    param($Row, [int]$Column, $Value)
    if ($Column -le 0) { return }
    $cell = $Row.Cells.Item(1, $Column)
    try {
        $cell.NumberFormat = "@"
        $cell.Value2 = if ($null -eq $Value) { "" } else { [string]$Value }
    }
    finally { Release-Com $cell }
}

function Set-Value {
    param($Row, [int]$Column, $Value)
    if ($Column -le 0) { return }
    $cell = $Row.Cells.Item(1, $Column)
    try { $cell.Value2 = $Value }
    finally { Release-Com $cell }
}

function Copy-PreviousRowTemplate {
    param($Table, $NewRow)
    if ($Table.ListRows.Count -le 1) { return }

    $previous = $null
    try {
        $previous = $Table.ListRows.Item($Table.ListRows.Count - 1).Range
        $previous.Copy()
        $NewRow.PasteSpecial(-4122) | Out-Null # formats
        $NewRow.PasteSpecial(6) | Out-Null     # validation

        for ($i = 1; $i -le $Table.ListColumns.Count; $i++) {
            $source = $null
            $target = $null
            try {
                $source = $previous.Cells.Item(1, $i)
                $target = $NewRow.Cells.Item(1, $i)
                if ($source.HasFormula) {
                    try { $target.Formula2R1C1 = $source.Formula2R1C1 }
                    catch { $target.FormulaR1C1 = $source.FormulaR1C1 }
                }
            }
            finally {
                Release-Com $target
                Release-Com $source
            }
        }
    }
    finally { Release-Com $previous }
}

function Build-Notes {
    param($Invoice, [bool]$IncludeNif)
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$Invoice.descripcio)) {
        $parts += "Descripcio: $(([string]$Invoice.descripcio).Trim())"
    }
    if ($IncludeNif -and -not [string]::IsNullOrWhiteSpace([string]$Invoice.nif)) {
        $parts += "NIF JSON: $(([string]$Invoice.nif).Trim().ToUpperInvariant())"
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Invoice.observacions)) {
        $parts += ([string]$Invoice.observacions).Trim()
    }
    return ($parts -join " | ")
}

try {
    if ([string]::IsNullOrWhiteSpace($JsonPath)) {
        $JsonPath = Pick-File "Selecciona el JSON de Factures JSON" "Fitxers JSON (*.json)|*.json"
    }
    if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
        $WorkbookPath = Pick-File "Selecciona el llibre de control fiscal" "Llibres Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
    }

    $JsonPath = (Resolve-Path -LiteralPath $JsonPath).Path
    $WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path

    $payload = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int]$payload.schema_version -ne 1) { throw "schema_version no compatible." }
    $invoices = @($payload.factures)
    if ($invoices.Count -eq 0) { throw "El JSON no conte factures." }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    if (-not $NoBackup -and -not $DryRun) {
        $folder = Split-Path -Parent $WorkbookPath
        $name = [IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
        $extension = [IO.Path]::GetExtension($WorkbookPath)
        $backup = Join-Path $folder "$name.backup_$timestamp$extension"
        Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force
        Write-Host "Copia de seguretat: $backup" -ForegroundColor DarkGray
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    if ($workbook.ReadOnly) { throw "El llibre esta obert o bloquejat en mode nomes lectura." }
    $worksheet = $workbook.Worksheets.Item($WorksheetName)
    $table = $worksheet.ListObjects.Item($TableName)

    $columns = @{
        EstatFiscal     = (Get-ColumnIndex $table @("Estat_Fiscal", "Incloure"))
        Tipus           = (Get-ColumnIndex $table @("Tipus"))
        Data            = (Get-ColumnIndex $table @("Data"))
        NumFactura      = (Get-ColumnIndex $table @("Num_Factura"))
        NomAlias        = (Get-ColumnIndex $table @("Nom_Alias"))
        NifInput        = (Get-ColumnIndex $table @("NIF_Entrada", "NIF_JSON", "NIF_Manual") -Optional)
        IvaPct          = (Get-ColumnIndex $table @("%IVA"))
        Bi              = (Get-ColumnIndex $table @("BI"))
        IrpfReper       = (Get-ColumnIndex $table @("%IRPF/%Reper", "%IRPF", "%Reper"))
        Observacions    = (Get-ColumnIndex $table @("Observacions"))
        PeriodeDeclarat = (Get-ColumnIndex $table @("Periode_Declarat") -Optional)
        Origen          = (Get-ColumnIndex $table @("Origen"))
        IdExtern        = (Get-ColumnIndex $table @("ID_Extern"))
    }

    $existingIds = @{}
    $existingInvoices = @{}
    $dataRange = $table.DataBodyRange
    if ($null -ne $dataRange) {
        try {
            for ($r = 1; $r -le $dataRange.Rows.Count; $r++) {
                $id = [string]$dataRange.Cells.Item($r, $columns.IdExtern).Value2
                if (-not [string]::IsNullOrWhiteSpace($id)) { $existingIds[$id.Trim()] = $true }

                $tipus = $dataRange.Cells.Item($r, $columns.Tipus).Value2
                $alias = $dataRange.Cells.Item($r, $columns.NomAlias).Value2
                $number = $dataRange.Cells.Item($r, $columns.NumFactura).Value2
                $date = $dataRange.Cells.Item($r, $columns.Data).Value2
                $bi = $dataRange.Cells.Item($r, $columns.Bi).Value2
                if ($null -ne $date -and $null -ne $bi -and -not [string]::IsNullOrWhiteSpace([string]$number)) {
                    try { $existingInvoices[(Invoice-Key $tipus $alias $number $date $bi)] = $true } catch { }
                }
            }
        }
        finally { Release-Com $dataRange }
    }

    $imported = @()
    $duplicatesById = @()
    $duplicatesByData = @()

    foreach ($invoice in $invoices) {
        foreach ($required in @("id_extern", "tipus", "data", "num_factura", "nom_alias", "bi")) {
            if ($null -eq $invoice.$required -or [string]::IsNullOrWhiteSpace([string]$invoice.$required)) {
                throw "Factura no valida: falta $required."
            }
        }

        $type = ([string]$invoice.tipus).Trim()
        if ($type -notin @("Emesa", "Rebuda")) { throw "Tipus no valid: $type" }
        $id = ([string]$invoice.id_extern).Trim()
        $key = Invoice-Key $type $invoice.nom_alias $invoice.num_factura $invoice.data $invoice.bi

        if ($existingIds.ContainsKey($id)) {
            $duplicatesById += [string]$invoice.num_factura
            continue
        }
        if ($existingInvoices.ContainsKey($key)) {
            $duplicatesByData += [string]$invoice.num_factura
            continue
        }

        if (-not $DryRun) {
            $listRow = $null
            $newRow = $null
            try {
                $listRow = $table.ListRows.Add()
                $newRow = $listRow.Range
                Copy-PreviousRowTemplate $table $newRow

                Set-Text $newRow $columns.EstatFiscal "Pendent"
                Set-Text $newRow $columns.Tipus $type
                Set-Value $newRow $columns.Data ([DateTime]::ParseExact((To-IsoDate $invoice.data), "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture).ToOADate())
                Set-Text $newRow $columns.NumFactura ([string]$invoice.num_factura).Trim()
                Set-Text $newRow $columns.NomAlias ([string]$invoice.nom_alias).Trim()
                if ($null -ne $columns.NifInput) { Set-Text $newRow $columns.NifInput ([string]$invoice.nif).Trim().ToUpperInvariant() }
                Set-Value $newRow $columns.IvaPct (To-Number $invoice.iva_pct)
                Set-Value $newRow $columns.Bi (To-Number $invoice.bi)

                $percentage = if ($type -eq "Emesa") { To-Number $invoice.irpf_pct } else { To-Number $invoice.afectacio_pct }
                Set-Value $newRow $columns.IrpfReper $percentage
                Set-Text $newRow $columns.Observacions (Build-Notes $invoice ($null -eq $columns.NifInput))
                if ($null -ne $columns.PeriodeDeclarat) { Set-Text $newRow $columns.PeriodeDeclarat "" }
                Set-Text $newRow $columns.Origen "JSON"
                Set-Text $newRow $columns.IdExtern $id
            }
            finally {
                Release-Com $newRow
                Release-Com $listRow
            }
        }

        $existingIds[$id] = $true
        $existingInvoices[$key] = $true
        $imported += "$type - $($invoice.num_factura) - $($invoice.nom_alias)"
    }

    if (-not $DryRun -and $imported.Count -gt 0) {
        $excel.CalculateFull()
        $workbook.Save()
    }

    $logPath = Join-Path (Split-Path -Parent $JsonPath) ("{0}.import_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($JsonPath), $timestamp)
    @(
        "Factures JSON -> Excel",
        "Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "JSON: $JsonPath",
        "Excel: $WorkbookPath",
        "Mode prova: $DryRun",
        "",
        "Importades: $($imported.Count)",
        ($imported | ForEach-Object { "  + $_" }),
        "",
        "Duplicades per ID: $($duplicatesById.Count)",
        ($duplicatesById | ForEach-Object { "  = $_" }),
        "",
        "Duplicades per dades: $($duplicatesByData.Count)",
        ($duplicatesByData | ForEach-Object { "  ~ $_" })
    ) | Set-Content -LiteralPath $logPath -Encoding UTF8

    Write-Host ""
    Write-Host "Importacio completada" -ForegroundColor Green
    Write-Host "  Importades: $($imported.Count)"
    Write-Host "  Duplicades per ID: $($duplicatesById.Count)"
    Write-Host "  Duplicades per dades: $($duplicatesByData.Count)"
    Write-Host "  Registre: $logPath" -ForegroundColor DarkGray
    if ($DryRun) { Write-Host "Mode prova: no s'ha modificat el llibre." -ForegroundColor Yellow }
}
catch {
    $exitCode = 1
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    Release-Com $table
    Release-Com $worksheet
    Release-Com $workbook
    Release-Com $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $exitCode
