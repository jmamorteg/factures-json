[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$JsonPath,

    [Parameter(Position = 1)]
    [string]$WorkbookPath,

    [string]$WorksheetName = "Entrada_Factures",
    [string]$TableName = "tblEntradaFactures",
    [switch]$NoBackup,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$excel = $null
$workbook = $null
$worksheet = $null
$table = $null
$exitCode = 0

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
            # Ignore cleanup errors.
        }
    }
}

function Select-InputFile {
    param(
        [string]$Title,
        [string]$Filter
    )

    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.CheckFileExists = $true

    try {
        $result = $dialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            throw "Operacion cancelada."
        }
        return $dialog.FileName
    }
    finally {
        $dialog.Dispose()
    }
}

function Get-ColumnIndex {
    param(
        [object]$Table,
        [string[]]$Names,
        [switch]$Optional
    )

    foreach ($name in $Names) {
        $column = $null
        try {
            $column = $Table.ListColumns.Item($name)
            return [int]$column.Index
        }
        catch {
            # Try the next accepted header.
        }
        finally {
            Release-ComObject $column
        }
    }

    if ($Optional) {
        return $null
    }

    throw "Falta una columna obligatoria: $($Names -join ' / ')"
}

function Get-JsonProperty {
    param(
        [object]$Object,
        [string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Convert-ToNumber {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture)
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($Value -is [ValueType]) {
        return [double]$Value
    }

    $number = 0.0
    if ([double]::TryParse(
        $text.Trim(),
        [Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )) {
        return $number
    }

    if ([double]::TryParse(
        $text.Trim(),
        [Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::CurrentCulture,
        [ref]$number
    )) {
        return $number
    }

    throw "Valor numerico no valido: $Value"
}

function Convert-ToIsoDate {
    param([object]$Value)

    if ($Value -is [DateTime]) {
        return $Value.ToString("yyyy-MM-dd")
    }

    if ($Value -is [ValueType]) {
        return [DateTime]::FromOADate([double]$Value).ToString("yyyy-MM-dd")
    }

    $text = [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture)
    $date = [DateTime]::MinValue

    if ([DateTime]::TryParseExact(
        $text.Trim(),
        "yyyy-MM-dd",
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None,
        [ref]$date
    )) {
        return $date.ToString("yyyy-MM-dd")
    }

    if ([DateTime]::TryParse(
        $text.Trim(),
        [Globalization.CultureInfo]::CurrentCulture,
        [Globalization.DateTimeStyles]::None,
        [ref]$date
    )) {
        return $date.ToString("yyyy-MM-dd")
    }

    throw "Fecha no valida: $Value"
}

function Normalize-KeyPart {
    param([object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return ([string]$Value).Trim().ToUpperInvariant() -replace '\s+', ''
}

function Get-InvoiceKey {
    param(
        [object]$Type,
        [object]$Alias,
        [object]$Number,
        [object]$Date,
        [object]$BaseAmount
    )

    $amount = Convert-ToNumber $BaseAmount
    $amountText = $amount.ToString("0.#####", [Globalization.CultureInfo]::InvariantCulture)

    return @(
        (Normalize-KeyPart $Type),
        (Normalize-KeyPart $Alias),
        (Normalize-KeyPart $Number),
        (Convert-ToIsoDate $Date),
        $amountText
    ) -join "|"
}

function Get-CellValue {
    param(
        [object]$Range,
        [int]$Row,
        [int]$Column
    )

    $cell = $null
    try {
        $cell = $Range.Cells.Item($Row, $Column)
        return $cell.Value2
    }
    finally {
        Release-ComObject $cell
    }
}

function Set-CellText {
    param(
        [object]$RowRange,
        [int]$Column,
        [object]$Value
    )

    if ($Column -le 0) {
        return
    }

    $cell = $null
    try {
        $cell = $RowRange.Cells.Item(1, $Column)
        $text = ""
        if ($null -ne $Value) {
            $text = [string]$Value
        }
        $cell.NumberFormat = "@"
        $cell.Value2 = $text
    }
    finally {
        Release-ComObject $cell
    }
}

function Set-CellValue {
    param(
        [object]$RowRange,
        [int]$Column,
        [object]$Value
    )

    if ($Column -le 0) {
        return
    }

    $cell = $null
    try {
        $cell = $RowRange.Cells.Item(1, $Column)
        if ($null -eq $Value) {
            $cell.ClearContents() | Out-Null
        }
        else {
            $cell.Value2 = $Value
        }
    }
    finally {
        Release-ComObject $cell
    }
}

function Build-Notes {
    param(
        [object]$Invoice,
        [bool]$IncludeNif
    )

    $parts = @()
    $description = Get-JsonProperty $Invoice "descripcio"
    $nif = Get-JsonProperty $Invoice "nif"
    $observations = Get-JsonProperty $Invoice "observacions"

    if (-not [string]::IsNullOrWhiteSpace([string]$description)) {
        $parts += "Descripcio: $(([string]$description).Trim())"
    }

    if ($IncludeNif -and -not [string]::IsNullOrWhiteSpace([string]$nif)) {
        $parts += "NIF JSON: $(([string]$nif).Trim().ToUpperInvariant())"
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$observations)) {
        $parts += ([string]$observations).Trim()
    }

    return $parts -join " | "
}

try {
    if ([string]::IsNullOrWhiteSpace($JsonPath)) {
        $JsonPath = Select-InputFile -Title "Selecciona el JSON de Factures JSON" -Filter "Archivos JSON (*.json)|*.json"
    }

    if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
        $WorkbookPath = Select-InputFile -Title "Selecciona el libro de control fiscal" -Filter "Libros Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
    }

    $JsonPath = (Resolve-Path -LiteralPath $JsonPath).Path
    $WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path

    $payload = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int](Get-JsonProperty $payload "schema_version") -ne 1) {
        throw "schema_version no compatible."
    }

    $invoices = @(Get-JsonProperty $payload "factures")
    if ($invoices.Count -eq 0) {
        throw "El JSON no contiene facturas."
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    if ((-not $NoBackup) -and (-not $DryRun)) {
        $folder = Split-Path -Parent $WorkbookPath
        $name = [IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
        $extension = [IO.Path]::GetExtension($WorkbookPath)
        $backupPath = Join-Path $folder "$name.backup_$timestamp$extension"
        Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
        Write-Host "Copia de seguridad: $backupPath" -ForegroundColor DarkGray
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    if ($workbook.ReadOnly) {
        throw "El libro esta abierto o bloqueado en modo solo lectura."
    }

    $worksheet = $workbook.Worksheets.Item($WorksheetName)
    $table = $worksheet.ListObjects.Item($TableName)

    $columns = @{
        State = Get-ColumnIndex $table @("Estat_Fiscal", "Incloure")
        Type = Get-ColumnIndex $table @("Tipus")
        Date = Get-ColumnIndex $table @("Data")
        Number = Get-ColumnIndex $table @("Num_Factura")
        Description = Get-ColumnIndex $table @("Descripcio", "Descripcio_Factura", "Descripcio factura") -Optional
        Alias = Get-ColumnIndex $table @("Nom_Alias")
        NifInput = Get-ColumnIndex $table @("NIF_Entrada", "NIF_JSON", "NIF_Manual") -Optional
        VatRate = Get-ColumnIndex $table @("%IVA")
        BaseAmount = Get-ColumnIndex $table @("BI")
        IrpfOrAllocation = Get-ColumnIndex $table @("%IRPF/%Reper", "%IRPF", "%Reper")
        Notes = Get-ColumnIndex $table @("Observacions")
        DeclaredPeriod = Get-ColumnIndex $table @("Periode_Declarat") -Optional
        Origin = Get-ColumnIndex $table @("Origen")
        ExternalId = Get-ColumnIndex $table @("ID_Extern")
    }

    $existingIds = @{}
    $existingKeys = @{}
    $dataBodyRange = $table.DataBodyRange

    if ($null -ne $dataBodyRange) {
        try {
            for ($rowNumber = 1; $rowNumber -le $dataBodyRange.Rows.Count; $rowNumber++) {
                $existingIdValue = Get-CellValue $dataBodyRange $rowNumber $columns.ExternalId
                $existingId = ""
                if ($null -ne $existingIdValue) {
                    $existingId = [string]$existingIdValue
                }

                if (-not [string]::IsNullOrWhiteSpace($existingId)) {
                    $existingIds[$existingId.Trim()] = $true
                }

                $existingType = Get-CellValue $dataBodyRange $rowNumber $columns.Type
                $existingAlias = Get-CellValue $dataBodyRange $rowNumber $columns.Alias
                $existingNumber = Get-CellValue $dataBodyRange $rowNumber $columns.Number
                $existingDate = Get-CellValue $dataBodyRange $rowNumber $columns.Date
                $existingBase = Get-CellValue $dataBodyRange $rowNumber $columns.BaseAmount

                if (($null -ne $existingDate) -and
                    ($null -ne $existingBase) -and
                    (-not [string]::IsNullOrWhiteSpace([string]$existingNumber))) {
                    try {
                        $existingKey = Get-InvoiceKey $existingType $existingAlias $existingNumber $existingDate $existingBase
                        $existingKeys[$existingKey] = $true
                    }
                    catch {
                        # Ignore incomplete historical rows.
                    }
                }
            }
        }
        finally {
            Release-ComObject $dataBodyRange
        }
    }

    $imported = @()
    $duplicatesById = @()
    $duplicatesByData = @()

    foreach ($invoice in $invoices) {
        foreach ($requiredName in @("id_extern", "tipus", "data", "num_factura", "nom_alias", "bi")) {
            $requiredValue = Get-JsonProperty $invoice $requiredName
            $requiredText = ""
            if ($null -ne $requiredValue) {
                $requiredText = [string]$requiredValue
            }

            if (($null -eq $requiredValue) -or [string]::IsNullOrWhiteSpace($requiredText)) {
                throw "Factura no valida: falta $requiredName."
            }
        }

        $type = ([string](Get-JsonProperty $invoice "tipus")).Trim()
        if ($type -notin @("Emesa", "Rebuda")) {
            throw "Tipus no valido: $type"
        }

        $externalId = ([string](Get-JsonProperty $invoice "id_extern")).Trim()
        $invoiceKey = Get-InvoiceKey $type (Get-JsonProperty $invoice "nom_alias") (Get-JsonProperty $invoice "num_factura") (Get-JsonProperty $invoice "data") (Get-JsonProperty $invoice "bi")

        if ($existingIds.ContainsKey($externalId)) {
            $duplicatesById += [string](Get-JsonProperty $invoice "num_factura")
            continue
        }

        if ($existingKeys.ContainsKey($invoiceKey)) {
            $duplicatesByData += [string](Get-JsonProperty $invoice "num_factura")
            continue
        }

        if (-not $DryRun) {
            $listRow = $null
            $rowRange = $null

            try {
                $listRow = $table.ListRows.Add()
                $rowRange = $listRow.Range

                Set-CellText $rowRange $columns.State "Pendent"
                Set-CellText $rowRange $columns.Type $type

                $invoiceDate = [DateTime]::ParseExact(
                    (Convert-ToIsoDate (Get-JsonProperty $invoice "data")),
                    "yyyy-MM-dd",
                    [Globalization.CultureInfo]::InvariantCulture
                )
                Set-CellValue $rowRange $columns.Date $invoiceDate.ToOADate()

                $numberText = ([string](Get-JsonProperty $invoice "num_factura")).Trim()
                $aliasText = ([string](Get-JsonProperty $invoice "nom_alias")).Trim()
                Set-CellText $rowRange $columns.Number $numberText
                Set-CellText $rowRange $columns.Alias $aliasText

                if ($null -ne $columns.Description) {
                    Set-CellText $rowRange $columns.Description (Get-JsonProperty $invoice "descripcio")
                }

                if ($null -ne $columns.NifInput) {
                    $nif = [string](Get-JsonProperty $invoice "nif")
                    Set-CellText $rowRange $columns.NifInput $nif.Trim().ToUpperInvariant()
                }

                Set-CellValue $rowRange $columns.VatRate (Convert-ToNumber (Get-JsonProperty $invoice "iva_pct"))
                Set-CellValue $rowRange $columns.BaseAmount (Convert-ToNumber (Get-JsonProperty $invoice "bi"))

                $irpfOrAllocation = $null
                if ($type -eq "Emesa") {
                    $irpfOrAllocation = Convert-ToNumber (Get-JsonProperty $invoice "irpf_pct")
                }
                else {
                    $irpfOrAllocation = Convert-ToNumber (Get-JsonProperty $invoice "afectacio_pct")
                }
                Set-CellValue $rowRange $columns.IrpfOrAllocation $irpfOrAllocation

                $includeNifInNotes = ($null -eq $columns.NifInput)
                Set-CellText $rowRange $columns.Notes (Build-Notes $invoice $includeNifInNotes)

                if ($null -ne $columns.DeclaredPeriod) {
                    Set-CellText $rowRange $columns.DeclaredPeriod ""
                }

                Set-CellText $rowRange $columns.Origin "JSON"
                Set-CellText $rowRange $columns.ExternalId $externalId
            }
            finally {
                Release-ComObject $rowRange
                Release-ComObject $listRow
            }
        }

        $existingIds[$externalId] = $true
        $existingKeys[$invoiceKey] = $true
        $imported += "$type - $([string](Get-JsonProperty $invoice 'num_factura')) - $([string](Get-JsonProperty $invoice 'nom_alias'))"
    }

    if ((-not $DryRun) -and ($imported.Count -gt 0)) {
        $excel.CalculateFull()
        $workbook.Save()
    }

    $logPath = Join-Path (Split-Path -Parent $JsonPath) ("{0}.import_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($JsonPath), $timestamp)

    @(
        "Factures JSON -> Excel",
        "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "JSON: $JsonPath",
        "Excel: $WorkbookPath",
        "Modo prueba: $DryRun",
        "",
        "Importadas: $($imported.Count)",
        ($imported | ForEach-Object { "  + $_" }),
        "",
        "Duplicadas por ID: $($duplicatesById.Count)",
        ($duplicatesById | ForEach-Object { "  = $_" }),
        "",
        "Duplicadas por datos: $($duplicatesByData.Count)",
        ($duplicatesByData | ForEach-Object { "  ~ $_" })
    ) | Set-Content -LiteralPath $logPath -Encoding UTF8

    Write-Host ""
    Write-Host "Importacion completada" -ForegroundColor Green
    Write-Host "  Importadas: $($imported.Count)"
    Write-Host "  Duplicadas por ID: $($duplicatesById.Count)"
    Write-Host "  Duplicadas por datos: $($duplicatesByData.Count)"
    Write-Host "  Registro: $logPath" -ForegroundColor DarkGray

    if ($DryRun) {
        Write-Host "Modo prueba: no se ha modificado el libro." -ForegroundColor Yellow
    }
}
catch {
    $exitCode = 1
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Tipo: $($_.Exception.GetType().FullName)" -ForegroundColor DarkGray

    if ($null -ne $_.InvocationInfo) {
        Write-Host "Linea: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
        Write-Host "Codigo: $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkGray
    }
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }

    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }

    Release-ComObject $table
    Release-ComObject $worksheet
    Release-ComObject $workbook
    Release-ComObject $excel

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $exitCode
