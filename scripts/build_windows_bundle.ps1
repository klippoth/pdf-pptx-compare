$ErrorActionPreference = "Stop"

function Invoke-NativeCommand {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Executable,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
  )

  & $Executable @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Command failed with exit code ${LASTEXITCODE}: $Executable $($Arguments -join ' ')"
  }
}

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Resolve-PythonCommand {
  $python = Get-Command python -ErrorAction SilentlyContinue
  if ($python) {
    return @{
      Executable = $python.Source
    }
  }

  $python3 = Get-Command python3 -ErrorAction SilentlyContinue
  if ($python3) {
    return @{
      Executable = $python3.Source
    }
  }

  $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
  if ($pyLauncher) {
    return @{
      Executable = $pyLauncher.Source
    }
  }

  throw "No Python 3 launcher was found. Install Python 3 for Windows and ensure 'py' or 'python' is on PATH."
}

function Assert-PathExists([string]$path, [string]$message) {
  if (-not (Test-Path $path)) {
    throw $message
  }
}

$venvPath = ".venv-windows"
$pythonExe = Join-Path $venvPath "Scripts\python.exe"
$pyInstallerExe = Join-Path $venvPath "Scripts\pyinstaller.exe"
$pythonCommand = Resolve-PythonCommand
$distPath = "dist-windows"
$workPath = "build-windows"
$specPath = Join-Path $workPath "spec"
$entryScript = Join-Path $root "launch_app.py"
$staticSource = Join-Path $root "app\static"

if (-not (Test-Path $pythonExe)) {
  & $pythonCommand.Executable -m venv $venvPath
  if ($LASTEXITCODE -ne 0) {
    throw "Command failed with exit code ${LASTEXITCODE}: $($pythonCommand.Executable) -m venv $venvPath"
  }
}

Assert-PathExists $pythonExe "Windows Python is not installed correctly. Install Python 3 for Windows from python.org, reopen PowerShell, and rerun this script."
Assert-PathExists $entryScript "Could not find launch_app.py in the project root."
Assert-PathExists $staticSource "Could not find app\static in the project root."

Invoke-NativeCommand $pythonExe -m pip install --upgrade pip
Invoke-NativeCommand $pythonExe -m pip install -r requirements.txt pyinstaller

New-Item -ItemType Directory -Force -Path $distPath | Out-Null
New-Item -ItemType Directory -Force -Path $workPath | Out-Null
New-Item -ItemType Directory -Force -Path $specPath | Out-Null

Invoke-NativeCommand $pyInstallerExe `
  --noconfirm `
  --clean `
  --onedir `
  --windowed `
  --distpath $distPath `
  --workpath $workPath `
  --specpath $specPath `
  --name "PDFtoPPTXReference" `
  --add-data "${staticSource};app/static" `
  $entryScript

$bundleRoot = Join-Path $distPath "PDFtoPPTXReference"
$localEnv = Join-Path $root ".env.local"
$localEnvExample = Join-Path $root ".env.local.example"
$includeLocalEnv = $env:PDF_PPTX_BUNDLE_LOCAL_ENV
$buildNoAi = $env:PDF_PPTX_BUILD_NO_AI
$shouldBundleLocalEnv = $true
if ($includeLocalEnv -and ($includeLocalEnv.ToLower() -eq "0" -or $includeLocalEnv.ToLower() -eq "false")) {
  $shouldBundleLocalEnv = $false
}
if ($buildNoAi -and ($buildNoAi.ToLower() -eq "1" -or $buildNoAi.ToLower() -eq "true")) {
  "PDF_PPTX_ENABLE_AI_QC=false`n" | Set-Content (Join-Path $bundleRoot ".env.local") -Encoding UTF8
  Write-Host "Bundled no-AI .env.local into the app package."
}
elseif ($shouldBundleLocalEnv -and (Test-Path $localEnv)) {
  Copy-Item $localEnv (Join-Path $bundleRoot ".env.local") -Force
  Write-Host "Bundled project-local .env.local into the app package."
}
if (Test-Path $localEnvExample) {
  Copy-Item $localEnvExample (Join-Path $bundleRoot ".env.local.example") -Force
}

Write-Host ""
Write-Host "Build complete."
Write-Host "Send the folder at dist-windows\PDFtoPPTXReference to the Windows user."
