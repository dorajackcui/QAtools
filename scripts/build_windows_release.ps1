[CmdletBinding()]
param(
    [string]$Version = "",
    [string]$PythonCommand = "python",
    [string]$InnoSetupCommand = "",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$tagRulesPath = Join-Path $projectRoot "phraseloom\tag_rules.toml"
$iconPath = Join-Path $projectRoot "packaging\QAtools.ico"
$installerScript = Join-Path $projectRoot "packaging\QAtools.iss"
$guiEntry = Join-Path $projectRoot "toolshub_gui.py"
$cliEntry = Join-Path $projectRoot "qatools_cli.py"
$originalTemp = [Environment]::GetEnvironmentVariable("TEMP", "Process")
$originalTmp = [Environment]::GetEnvironmentVariable("TMP", "Process")
$originalPath = [Environment]::GetEnvironmentVariable("PATH", "Process")

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

function Resolve-InnoSetupCompiler {
    if ($InnoSetupCommand) {
        $requestedCompiler = Get-Command $InnoSetupCommand -ErrorAction SilentlyContinue
        if ($requestedCompiler) {
            return $requestedCompiler.Source
        }
        if (Test-Path -LiteralPath $InnoSetupCommand -PathType Leaf) {
            return (Resolve-Path -LiteralPath $InnoSetupCommand).Path
        }
        throw "Inno Setup compiler not found: $InnoSetupCommand"
    }

    $availableCompiler = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
    if ($availableCompiler) {
        return $availableCompiler.Source
    }
    foreach ($candidate in @(
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }
    throw "Inno Setup 6 is required to build the Windows installer."
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
    # Put Windows system DLLs ahead of unrelated native toolchains that may be
    # injected into PATH by the build host. Otherwise PyInstaller can bundle a
    # third-party DLL with the same name as a Windows DLL (for example ICU),
    # leaving the frozen Qt runtime unable to load.
    $system32Path = Join-Path $env:SystemRoot "System32"
    $env:PATH = "$system32Path;$originalPath"

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
    $innoSetupCompiler = Resolve-InnoSetupCompiler

    if (-not $SkipTests) {
        Invoke-ProjectPython `
            -Description "Regression tests" `
            -Arguments @("-m", "unittest", "discover", "-s", "tests", "-v")
    }

    $installerName = "QAtools-v$Version-windows-$architecture-setup"
    $legacyReleaseName = "QAtools-v$Version-windows-$architecture"
    $distRoot = Join-Path $projectRoot "dist"
    $installerPath = Join-Path $distRoot "$installerName.exe"
    $legacyReleaseDir = Join-Path $distRoot $legacyReleaseName
    $legacyZipPath = Join-Path $distRoot "$legacyReleaseName.zip"
    $buildRoot = Join-Path $projectRoot "build\windows-release"
    $appDir = Join-Path $buildRoot "installer-app"
    $exeDir = Join-Path $buildRoot "executables"
    $workDir = Join-Path $buildRoot "work"
    $specDir = Join-Path $buildRoot "spec"

    foreach ($path in @($installerPath, $legacyReleaseDir, $legacyZipPath, $buildRoot)) {
        Remove-ProjectArtifact -Path $path
    }
    New-Item -ItemType Directory -Path $distRoot, $appDir, $exeDir, $workDir, $specDir -Force | Out-Null

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
        "--icon", $iconPath,
        "--paths", $projectRoot,
        "--distpath", $exeDir,
        "--workpath", $workDir,
        "--specpath", $specDir,
        "--add-data", "$tagRulesPath;phraseloom"
    )

    Invoke-ProjectPython `
        -Description "GUI build" `
        -Arguments ($commonArguments + @(
            "--onedir",
            "--contents-directory", "_internal",
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
        "--onefile",
        "--console",
        "--name", "QAtools-CLI"
    )
    foreach ($module in $hiddenImports) {
        $cliArguments += @("--hidden-import", $module)
    }
    $cliArguments += $cliEntry

    Invoke-ProjectPython -Description "CLI build" -Arguments $cliArguments

    $guiBundleDir = Join-Path $exeDir "QAtools"
    Copy-Item -LiteralPath (Join-Path $guiBundleDir "QAtools.exe") -Destination $appDir
    Copy-Item -LiteralPath (Join-Path $guiBundleDir "_internal") -Destination $appDir -Recurse
    Copy-Item -LiteralPath (Join-Path $exeDir "QAtools-CLI.exe") -Destination $appDir
    Copy-Item -LiteralPath (Join-Path $projectRoot "packaging\QAtools-CLI.cmd") -Destination $appDir
    Copy-Item -LiteralPath (Join-Path $projectRoot "packaging\README-Windows.txt") -Destination $appDir

    # Verify the frozen programs, including every persistent Qt page, without
    # showing the application window.
    $smokeTemp = Join-Path $buildRoot "smoke-temp"
    New-Item -ItemType Directory -Path $smokeTemp -Force | Out-Null
    $env:TEMP = $smokeTemp
    $env:TMP = $smokeTemp

    & (Join-Path $appDir "QAtools-CLI.exe") --version
    if ($LASTEXITCODE -ne 0) {
        throw "Frozen CLI smoke test failed with exit code $LASTEXITCODE"
    }
    $guiSmokeProcess = Start-Process `
        -FilePath (Join-Path $appDir "QAtools.exe") `
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
    ) | Set-Content -LiteralPath (Join-Path $appDir "VERSION.txt") -Encoding UTF8

    $hashLines = Get-ChildItem -LiteralPath $appDir -Filter "*.exe" |
        Sort-Object Name |
        ForEach-Object {
            $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            "$hash  $($_.Name)"
        }
    $hashLines | Set-Content -LiteralPath (Join-Path $appDir "SHA256SUMS.txt") -Encoding ASCII

    & $innoSetupCompiler `
        "/DAppVersion=$Version" `
        "/DSourceDir=$appDir" `
        "/O$distRoot" `
        "/F$installerName" `
        $installerScript
    if ($LASTEXITCODE -ne 0) {
        throw "Windows installer build failed with exit code $LASTEXITCODE"
    }
    if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
        throw "Windows installer was not created: $installerPath"
    }

    Write-Host ""
    Write-Host "Windows installer created:"
    Write-Host "  Installer: $installerPath"
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
    if ($null -eq $originalPath) {
        Remove-Item Env:PATH -ErrorAction SilentlyContinue
    }
    else {
        $env:PATH = $originalPath
    }
    Pop-Location
}
