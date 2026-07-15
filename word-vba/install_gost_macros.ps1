param(
    [string]$MacroPath = (Join-Path $PSScriptRoot "GostFormatting.bas"),
    [string]$ModuleName = "GostFormatting",
    [string]$BackupDir = (Join-Path $PSScriptRoot "backups"),
    [string[]]$LegacyProcedureNames = @(
        "РисункиФорматирование",
        "ТаблицыФорматирование",
        "ОбновитьНумерациюЗаголовков",
        "AlignTablesAndFigures",
        "UpdateGostHeadingNumbering"
    )
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $MacroPath)) {
    throw "Macro file not found: $MacroPath"
}

$resolvedMacroPath = (Resolve-Path -LiteralPath $MacroPath).Path
$backupPaths = @()
$word = $null
$templateDocument = $null

function Test-ComponentContainsProcedure {
    param(
        [Parameter(Mandatory = $true)]$Component,
        [Parameter(Mandatory = $true)][string[]]$ProcedureNames
    )

    if ($Component.Type -ne 1) {
        return $false
    }

    $codeModule = $Component.CodeModule
    if ($codeModule.CountOfLines -le 0) {
        return $false
    }

    $source = $codeModule.Lines(1, $codeModule.CountOfLines)
    foreach ($procedureName in $ProcedureNames) {
        $pattern = "(?im)^\s*(Public\s+|Private\s+)?Sub\s+$([regex]::Escape($procedureName))\s*\("
        if ($source -match $pattern) {
            return $true
        }
    }

    return $false
}

try {
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $template = $word.NormalTemplate
    $templateDocument = $template.OpenAsDocument()
    $project = $templateDocument.VBProject

    for ($index = $project.VBComponents.Count; $index -ge 1; $index--) {
        $component = $project.VBComponents.Item($index)
        if ($component.Name -eq $ModuleName -or (Test-ComponentContainsProcedure -Component $component -ProcedureNames $LegacyProcedureNames)) {
            $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
            $backupPath = Join-Path $BackupDir "$($component.Name)-$timestamp.bas"
            $component.Export($backupPath)
            $backupPaths += $backupPath
            $project.VBComponents.Remove($component)
        }
    }

    $beforeImportCount = [int]$project.VBComponents.Count
    $imported = $project.VBComponents.Import($resolvedMacroPath)
    if ($null -eq $imported) {
        $afterImportCount = [int]$project.VBComponents.Count
        if ($afterImportCount -gt $beforeImportCount) {
            $imported = $project.VBComponents.Item($afterImportCount)
        }
    }
    if ($null -ne $imported) {
        $imported.Name = $ModuleName
    }
    elseif ($project.VBComponents.Count -eq $beforeImportCount) {
        throw "Word did not import the VBA module."
    }
    $templateDocument.Save()

    [pscustomobject]@{
        Template = $templateDocument.FullName
        Imported = $resolvedMacroPath
        Backups = $backupPaths
    } | ConvertTo-Json -Depth 3
}
catch {
    throw "Не удалось импортировать VBA-модуль в Word. Проверьте, что Word установлен и включена настройка 'Trust access to the VBA project object model'. Исходная ошибка: $($_.Exception.Message)"
}
finally {
    if ($templateDocument -ne $null) {
        try {
            $templateDocument.Close($false)
        }
        catch {
        }
    }
    if ($word -ne $null) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
