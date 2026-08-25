[CmdletBinding()]
param(
    [string]$Version = "",
    [string]$PythonCommand = "python",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$tagRulesPath = Join-Path $projectRoot "phraseloom\tag_rules.toml"
$guiEntry = Join-Path $projectRoot "toolshub_gui.py"
$cliEntry = Join-Path $projectRoot "qatools_cli.py"
$originalTemp = [Environment]::GetEnvironmentVariable("TEMP", "Process")
$originalTmp = [Environment]::GetEnvironmentVariable("TMP", "Process")

function Invoke-ProjectPython {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,
        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    & $PythonCommand @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE"
    }
}

function Remove-ProjectArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $resolvedCandidate = [System.IO.Path]::GetFullPath($Path)
    $rootPrefix = $projectRoot.TrimEnd("\", "/") + [System.IO.Path]::DirectorySeparatorChar
    if (-not $resolvedCandidate.StartsWith(
        $rootPrefix,
        [System.StringComparison]::OrdinalIgnoreCase
    )) {
        throw "Refusing to remove a path outside the project: $resolvedCandidate"
    }
    if (Test-Path -LiteralPath $resolvedCandidate) {
        Remove-Item -LiteralPath $resolvedCandidate -Recurse -Force
    }
}

Push-Location $projectRoot
try {
    if (-not $Version) {
        $versionOutput = & $PythonCommand -c "from qatools import __version__; print(__version__)"
        if ($LASTEXITCODE -ne 0) {
            throw "Could not read the QAtools version."
        }
        $Version = ($versionOutput | Select-Object -Last 1).Trim()
    }
    if ($Version -notmatch '^[0-9A-Za-z][0-9A-Za-z._-]*$') {
        throw "The version contains characters that are unsafe in a file name: $Version"
    }

    $pointerBits = & $PythonCommand -c "import struct; print(struct.calcsize('P') * 8)"
    if ($LASTEXITCODE -ne 0) {
        throw "Could not determine the Python architecture."
    }
    $architecture = if (($pointerBits | Select-Object -Last 1).Trim() -eq "64") {
        "x64"
    }
    else {
        "x86"
    }

    Invoke-ProjectPython `
        -Description "PyInstaller check" `
        -Arguments @("-m", "PyInstaller", "--version")

    if (-not $SkipTests) {
        Invoke-ProjectPython `
            -Description "Regression tests" `
            -Arguments @("-m", "unittest", "discover", "-s", "tests", "-v")
    }

    $releaseName = "QAtools-v$Version-windows-$architecture"
    $distRoot = Join-Path $projectRoot "dist"
    $releaseDir = Join-Path $distRoot $releaseName
    $zipPath = Join-Path $distRoot "$releaseName.zip"
    $buildRoot = Join-Path $projectRoot "build\windows-release"
    $exeDir = Join-Path $buildRoot "executables"
    $workDir = Join-Path $buildRoot "work"
    $specDir = Join-Path $buildRoot "spec"

    foreach ($path in @($releaseDir, $zipPath, $buildRoot)) {
        Remove-ProjectArtifact -Path $path
    }
    New-Item -ItemType Directory -Path $releaseDir, $exeDir, $workDir, $specDir -Force | Out-Null

    Invoke-ProjectPython `
        -Description "PySide6 runtime check" `
        -Arguments @(
            "-c",
            "from PySide6 import QtCore; print(QtCore.__version__)"
        )

    $commonArguments = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--paths", $projectRoot,
        "--distpath", $exeDir,
        "--workpath", $workDir,
        "--specpath", $specDir,
        "--add-data", "$tagRulesPath;phraseloom"
    )

    Invoke-ProjectPython `
        -Description "GUI build" `
        -Arguments ($commonArguments + @(
            "--windowed",
            "--name", "QAtools",
            $guiEntry
        ))

    $hiddenImports = @(
        "toolshub_gui",
        "tools.workflow.cli",
        "phraseloom.cli",
        "tools.term_pair_checker.extract_terms_from_excel",
        "tools.tag_placeholder_checker.check_tags_and_placeholders",
        "tools.line_break_checker.check_line_breaks",
        "tools.source_consistency_checker.check_source_consistency",
        "tools.chinese_target_checker.check_chinese_target",
        "tools.french_nbsp_restorer.restore_french_nbsp",
        "tools.excel_batcher.excel_batcher",
        "tools.excel_merger.merge_active_sheets",
        "tools.xbench_report_transformer.transform_xbench_report"
    )
    $cliArguments = $commonArguments + @(
        "--console",
        "--name", "QAtools-CLI"
    )
    foreach ($module in $hiddenImports) {
        $cliArguments += @("--hidden-import", $module)
    }
    $cliArguments += $cliEntry

    Invoke-ProjectPython -Description "CLI build" -Arguments $cliArguments

    Copy-Item -LiteralPath (Join-Path $exeDir "QAtools.exe") -Destination $releaseDir
    Copy-Item -LiteralPath (Join-Path $exeDir "QAtools-CLI.exe") -Destination $releaseDir
    Copy-Item -LiteralPath (Join-Path $projectRoot "packaging\QAtools-CLI.cmd") -Destination $releaseDir
    Copy-Item -LiteralPath (Join-Path $projectRoot "packaging\README-Windows.txt") -Destination $releaseDir

    # Verify the frozen programs, including every persistent Qt page, without
    # showing the application window.
    $smokeTemp = Join-Path $buildRoot "smoke-temp"
    New-Item -ItemType Directory -Path $smokeTemp -Force | Out-Null
    $env:TEMP = $smokeTemp
    $env:TMP = $smokeTemp

    & (Join-Path $releaseDir "QAtools-CLI.exe") --version
    if ($LASTEXITCODE -ne 0) {
        throw "Frozen CLI smoke test failed with exit code $LASTEXITCODE"
    }
    $guiSmokeProcess = Start-Process `
        -FilePath (Join-Path $releaseDir "QAtools.exe") `
        -ArgumentList "--smoke-test" `
        -WindowStyle Hidden `
        -Wait `
        -PassThru
    if ($guiSmokeProcess.ExitCode -ne 0) {
        throw "Frozen GUI smoke test failed with exit code $($guiSmokeProcess.ExitCode)"
    }

    $builtAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss K"
    @(
        "QAtools $Version",
        "Windows $architecture",
        "Built at $builtAt",
        "Python $((& $PythonCommand --version) -replace '^Python\s+', '')"
    ) | Set-Content -LiteralPath (Join-Path $releaseDir "VERSION.txt") -Encoding UTF8

    $hashLines = Get-ChildItem -LiteralPath $releaseDir -Filter "*.exe" |
        Sort-Object Name |
        ForEach-Object {
            $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            "$hash  $($_.Name)"
        }
    $hashLines | Set-Content -LiteralPath (Join-Path $releaseDir "SHA256SUMS.txt") -Encoding ASCII

    Compress-Archive `
        -Path (Join-Path $releaseDir "*") `
        -DestinationPath $zipPath `
        -CompressionLevel Optimal

    Write-Host ""
    Write-Host "Windows portable release created:"
    Write-Host "  Directory: $releaseDir"
    Write-Host "  Archive:   $zipPath"
}
finally {
    if ($null -eq $originalTemp) {
        Remove-Item Env:TEMP -ErrorAction SilentlyContinue
    }
    else {
        $env:TEMP = $originalTemp
    }
    if ($null -eq $originalTmp) {
        Remove-Item Env:TMP -ErrorAction SilentlyContinue
    }
    else {
        $env:TMP = $originalTmp
    }
    Pop-Location
}
