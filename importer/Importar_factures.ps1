[CmdletBinding()]
param(
    [Parameter(Position=0)][string]$JsonPath,
    [Parameter(Position=1)][string]$WorkbookPath,
    [string]$WorksheetName="Entrada_Factures",
    [string]$TableName="tblEntradaFactures",
    [switch]$NoBackup,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
$excel=$null; $book=$null; $sheet=$null; $table=$null; $exitCode=0

function Release-Com($o) {
    if ($null -ne $o) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch {}
    }
}

function Pick-File([string]$title,[string]$filter) {
    Add-Type -AssemblyName System.Windows.Forms
    $d=New-Object System.Windows.Forms.OpenFileDialog
    $d.Title=$title; $d.Filter=$filter; $d.CheckFileExists=$true
    try {
        if ($d.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            throw "Operacio cancel-lada."
        }
        return $d.FileName
    } finally { $d.Dispose() }
}

function Col($table,[string[]]$names,[switch]$optional) {
    foreach($name in $names) {
        $c=$null
        try { $c=$table.ListColumns.Item($name); return [int]$c.Index }
        catch {}
        finally { Release-Com $c }
    }
    if($optional){ return $null }
    throw "Falta una columna obligatoria: $($names -join ' / ')"
}

function Prop($object,[string]$name) {
    $p=$object.PSObject.Properties[$name]
    if($null -eq $p){ return $null }
    return $p.Value
}

function Num($value) {
    if($null -eq $value -or [string]::IsNullOrWhiteSpace(([string]$value))){ return $null }
    if($value -is [ValueType]){ return [double]$value }
    $n=0.0; $s=([string]$value).Trim()
    if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){ return $n }
    if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::CurrentCulture,[ref]$n)){ return $n }
    throw "Valor numeric no valid: $value"
}

function Iso-Date($value) {
    if($value -is [DateTime]){ return $value.ToString("yyyy-MM-dd") }
    if($value -is [ValueType]){ return [DateTime]::FromOADate([double]$value).ToString("yyyy-MM-dd") }
    $d=[DateTime]::MinValue; $s=([string]$value).Trim()
    if([DateTime]::TryParseExact($s,"yyyy-MM-dd",[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$d)){ return $d.ToString("yyyy-MM-dd") }
    if([DateTime]::TryParse($s,[Globalization.CultureInfo]::CurrentCulture,[Globalization.DateTimeStyles]::None,[ref]$d)){ return $d.ToString("yyyy-MM-dd") }
    throw "Data no valida: $value"
}

function Norm($value) {
    if($null -eq $value){ return "" }
    return (([string]$value).Trim().ToUpperInvariant() -replace '\s+','')
}

function Invoice-Key($type,$alias,$number,$date,$bi) {
    $amount=(Num $bi).ToString("0.#####",[Globalization.CultureInfo]::InvariantCulture)
    return "$(Norm $type)|$(Norm $alias)|$(Norm $number)|$(Iso-Date $date)|$amount"
}

function Set-Text($row,[int]$column,$value) {
    if($column -le 0){ return }
    $cell=$null
    try {
        $cell=$row.Cells.Item(1,$column)
        $text=if($null -eq $value){""}else{[string]$value}
        $cell.NumberFormat="@"
        $cell.Value2=$text
    } finally { Release-Com $cell }
}

function Set-Value($row,[int]$column,$value) {
    if($column -le 0){ return }
    $cell=$null
    try { $cell=$row.Cells.Item(1,$column); $cell.Value2=$value }
    finally { Release-Com $cell }
}

function Notes($invoice,[bool]$includeNif) {
    $parts=@()
    $description=Prop $invoice "descripcio"
    $nif=Prop $invoice "nif"
    $observations=Prop $invoice "observacions"
    if(-not [string]::IsNullOrWhiteSpace(([string]$description))){ $parts+="Descripcio: $(([string]$description).Trim())" }
    if($includeNif -and -not [string]::IsNullOrWhiteSpace(([string]$nif))){ $parts+="NIF JSON: $(([string]$nif).Trim().ToUpperInvariant())" }
    if(-not [string]::IsNullOrWhiteSpace(([string]$observations))){ $parts+=([string]$observations).Trim() }
    return $parts -join " | "
}

try {
    if([string]::IsNullOrWhiteSpace($JsonPath)){
        $JsonPath=Pick-File "Selecciona el JSON de Factures JSON" "Fitxers JSON (*.json)|*.json"
    }
    if([string]::IsNullOrWhiteSpace($WorkbookPath)){
        $WorkbookPath=Pick-File "Selecciona el llibre de control fiscal" "Llibres Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
    }

    $JsonPath=(Resolve-Path -LiteralPath $JsonPath).Path
    $WorkbookPath=(Resolve-Path -LiteralPath $WorkbookPath).Path
    $payload=Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if([int]$payload.schema_version -ne 1){ throw "schema_version no compatible." }
    $invoices=@($payload.factures)
    if($invoices.Count -eq 0){ throw "El JSON no conte factures." }

    $stamp=Get-Date -Format "yyyyMMdd_HHmmss"
    if(-not $NoBackup -and -not $DryRun){
        $folder=Split-Path -Parent $WorkbookPath
        $name=[IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
        $extension=[IO.Path]::GetExtension($WorkbookPath)
        $backup=Join-Path $folder "$name.backup_$stamp$extension"
        Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force
        Write-Host "Copia de seguretat: $backup" -ForegroundColor DarkGray
    }

    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    $excel.ScreenUpdating=$false; $excel.EnableEvents=$false
    $book=$excel.Workbooks.Open($WorkbookPath,0,$false)
    if($book.ReadOnly){ throw "El llibre esta obert o bloquejat en mode nomes lectura." }
    $sheet=$book.Worksheets.Item($WorksheetName)
    $table=$sheet.ListObjects.Item($TableName)

    $cols=@{
        State=Col $table @("Estat_Fiscal","Incloure")
        Type=Col $table @("Tipus")
        Date=Col $table @("Data")
        Number=Col $table @("Num_Factura")
        Alias=Col $table @("Nom_Alias")
        Nif=Col $table @("NIF_Entrada","NIF_JSON","NIF_Manual") -optional
        Vat=Col $table @("%IVA")
        Base=Col $table @("BI")
        Irpf=Col $table @("%IRPF/%Reper","%IRPF","%Reper")
        Notes=Col $table @("Observacions")
        Period=Col $table @("Periode_Declarat") -optional
        Origin=Col $table @("Origen")
        Id=Col $table @("ID_Extern")
    }

    $ids=@{}; $keys=@{}
    $range=$table.DataBodyRange
    if($null -ne $range){
        try {
            for($r=1;$r -le $range.Rows.Count;$r++){
                $idValue=$range.Cells.Item($r,$cols.Id).Value2
                $id=if($null -eq $idValue){""}else{[string]$idValue}
                if(-not [string]::IsNullOrWhiteSpace($id)){ $ids[$id.Trim()]=$true }

                $type=$range.Cells.Item($r,$cols.Type).Value2
                $alias=$range.Cells.Item($r,$cols.Alias).Value2
                $number=$range.Cells.Item($r,$cols.Number).Value2
                $date=$range.Cells.Item($r,$cols.Date).Value2
                $bi=$range.Cells.Item($r,$cols.Base).Value2
                if($null -ne $date -and $null -ne $bi -and -not [string]::IsNullOrWhiteSpace(([string]$number))){
                    try { $keys[(Invoice-Key $type $alias $number $date $bi)]=$true } catch {}
                }
            }
        } finally { Release-Com $range }
    }

    $imported=@(); $duplicateIds=@(); $duplicateData=@()
    foreach($invoice in $invoices){
        foreach($required in @("id_extern","tipus","data","num_factura","nom_alias","bi")){
            $value=Prop $invoice $required
            if($null -eq $value -or [string]::IsNullOrWhiteSpace(([string]$value)){
                throw "Factura no valida: falta $required."
            }
        }

        $type=([string](Prop $invoice "tipus")).Trim()
        if($type -notin @("Emesa","Rebuda")){ throw "Tipus no valid: $type" }
        $id=([string](Prop $invoice "id_extern")).Trim()
        $key=Invoice-Key $type (Prop $invoice "nom_alias") (Prop $invoice "num_factura") (Prop $invoice "data") (Prop $invoice "bi")

        if($ids.ContainsKey($id)){ $duplicateIds+=[string](Prop $invoice "num_factura"); continue }
        if($keys.ContainsKey($key)){ $duplicateData+=[string](Prop $invoice "num_factura"); continue }

        if(-not $DryRun){
            $listRow=$null; $row=$null
            try {
                # ListRows.Add already propagates calculated-column formulas,
                # formats and validation in a normal Excel table.
                $listRow=$table.ListRows.Add()
                $row=$listRow.Range

                Set-Text $row $cols.State "Pendent"
                Set-Text $row $cols.Type $type
                $date=[DateTime]::ParseExact((Iso-Date (Prop $invoice "data")),"yyyy-MM-dd",[Globalization.CultureInfo]::InvariantCulture)
                Set-Value $row $cols.Date $date.ToOADate()
                Set-Text $row $cols.Number ([string](Prop $invoice "num_factura")).Trim()
                Set-Text $row $cols.Alias ([string](Prop $invoice "nom_alias")).Trim()

                if($null -ne $cols.Nif){
                    Set-Text $row $cols.Nif ([string](Prop $invoice "nif")).Trim().ToUpperInvariant()
                }

                Set-Value $row $cols.Vat (Num (Prop $invoice "iva_pct"))
                Set-Value $row $cols.Base (Num (Prop $invoice "bi"))
                $percentage=if($type -eq "Emesa"){ Num (Prop $invoice "irpf_pct") }else{ Num (Prop $invoice "afectacio_pct") }
                Set-Value $row $cols.Irpf $percentage
                Set-Text $row $cols.Notes (Notes $invoice ($null -eq $cols.Nif))
                if($null -ne $cols.Period){ Set-Text $row $cols.Period "" }
                Set-Text $row $cols.Origin "JSON"
                Set-Text $row $cols.Id $id
            } finally {
                Release-Com $row
                Release-Com $listRow
            }
        }

        $ids[$id]=$true; $keys[$key]=$true
        $imported+="$type - $([string](Prop $invoice 'num_factura')) - $([string](Prop $invoice 'nom_alias'))"
    }

    if(-not $DryRun -and $imported.Count -gt 0){
        $excel.CalculateFull()
        $book.Save()
    }

    $log=Join-Path (Split-Path -Parent $JsonPath) ("{0}.import_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($JsonPath),$stamp)
    @(
        "Factures JSON -> Excel",
        "Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "JSON: $JsonPath",
        "Excel: $WorkbookPath",
        "Mode prova: $DryRun",
        "",
        "Importades: $($imported.Count)",
        ($imported | ForEach-Object{"  + $_"}),
        "",
        "Duplicades per ID: $($duplicateIds.Count)",
        ($duplicateIds | ForEach-Object{"  = $_"}),
        "",
        "Duplicades per dades: $($duplicateData.Count)",
        ($duplicateData | ForEach-Object{"  ~ $_"})
    ) | Set-Content -LiteralPath $log -Encoding UTF8

    Write-Host ""
    Write-Host "Importacio completada" -ForegroundColor Green
    Write-Host "  Importades: $($imported.Count)"
    Write-Host "  Duplicades per ID: $($duplicateIds.Count)"
    Write-Host "  Duplicades per dades: $($duplicateData.Count)"
    Write-Host "  Registre: $log" -ForegroundColor DarkGray
    if($DryRun){ Write-Host "Mode prova: no s'ha modificat el llibre." -ForegroundColor Yellow }
}
catch {
    $exitCode=1
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Tipus: $($_.Exception.GetType().FullName)" -ForegroundColor DarkGray
    Write-Host "Linia: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    if(-not [string]::IsNullOrWhiteSpace($_.InvocationInfo.Line)){
        Write-Host "Ordre: $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkGray
    }
    if(-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)){
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    }
}
finally {
    if($null -ne $book){ try{$book.Close($false)}catch{} }
    if($null -ne $excel){ try{$excel.Quit()}catch{} }
    Release-Com $table; Release-Com $sheet; Release-Com $book; Release-Com $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

exit $exitCode
