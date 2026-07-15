[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$wordVbaRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $wordVbaRoot "GostFormatting.bas"
$installerPath = Join-Path $wordVbaRoot "install_gost_macros.ps1"
$metadataPath = Join-Path $wordVbaRoot "legacy\SOURCE_METADATA.json"
$source = [System.IO.File]::ReadAllText($modulePath)
$failures = [System.Collections.Generic.List[string]]::new()
$checks = 0

function Assert-SourceMatch {
    param(
        [Parameter(Mandatory = $true)][string]$Pattern,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $script:checks++
    if (-not [regex]::IsMatch($source, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline)) {
        $failures.Add($Description)
    }
}

Assert-SourceMatch -Pattern '^Option Explicit\s*$' -Description 'Option Explicit is missing.'

$publicProcedures = @(
    'РисункиФорматирование',
    'ТаблицыФорматирование',
    'ОбновитьНумерациюЗаголовков',
    'AlignTablesAndFigures',
    'UpdateGostHeadingNumbering'
)
foreach ($procedure in $publicProcedures) {
    Assert-SourceMatch -Pattern ("^Public\s+Sub\s+{0}\s*\(\s*\)" -f [regex]::Escape($procedure)) -Description "Public procedure is missing: $procedure"
}

Assert-SourceMatch -Pattern 'tableItem\.PreferredWidthType\s*=\s*wdPreferredWidthPercent' -Description 'Table width type is not percentage based.'
Assert-SourceMatch -Pattern 'tableItem\.PreferredWidth\s*=\s*100' -Description 'Tables are not configured to use the full available width.'
Assert-SourceMatch -Pattern 'tableItem\.Rows\(1\)\.HeadingFormat\s*=\s*True' -Description 'The first table row is not repeated as a heading.'
Assert-SourceMatch -Pattern 'tableItem\.Range\.NoProofing\s*=\s*True' -Description 'Table proofing is not disabled.'
Assert-SourceMatch -Pattern 'If\s+inlinePicture\.Width\s*>\s*maxWidth\s+Then' -Description 'Figure downscaling guard is missing.'
Assert-SourceMatch -Pattern 'HasAdjacentCaption\(' -Description 'Adjacent caption detection is missing.'
Assert-SourceMatch -Pattern 'Selection\.Range\.NoProofing\s*=\s*True' -Description 'Caption proofing is not disabled.'

$checks++
if ([regex]::IsMatch($source, 'If\s+inlinePicture\.Width\s*<\s*maxWidth\s+Then', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
    $failures.Add('The module contains an image-upscaling branch.')
}

$tokens = $null
$parseErrors = $null
[System.Management.Automation.Language.Parser]::ParseFile($installerPath, [ref]$tokens, [ref]$parseErrors) | Out-Null
$checks++
if ($parseErrors.Count -gt 0) {
    $failures.Add("Installer parser errors: $($parseErrors.Message -join '; ')")
}

$metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
foreach ($original in $metadata.originals) {
    $legacyPath = Join-Path (Join-Path $wordVbaRoot 'legacy') $original.fileName
    $checks++
    if (-not (Test-Path -LiteralPath $legacyPath -PathType Leaf)) {
        $failures.Add("Historical source is missing: $($original.fileName)")
        continue
    }

    $actualFile = Get-Item -LiteralPath $legacyPath
    $actualHash = (Get-FileHash -LiteralPath $legacyPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $checks += 2
    if ($actualFile.Length -ne [long]$original.length) {
        $failures.Add("Historical source length differs: $($original.fileName)")
    }
    if ($actualHash -ne [string]$original.sha256) {
        $failures.Add("Historical source hash differs: $($original.fileName)")
    }
}

if ($failures.Count -gt 0) {
    throw "VBA static checks failed ($($failures.Count)/$checks):`n- $($failures -join "`n- ")"
}

[pscustomobject]@{
    Status = 'passed'
    Checks = $checks
    Module = $modulePath
    HistoricalSources = $metadata.originals.Count
} | ConvertTo-Json -Depth 3
