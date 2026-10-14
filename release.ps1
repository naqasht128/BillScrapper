# Bill Scrapper - Quick PowerShell Release Shortcut
param (
    [string]$Version = "",
    [string]$Message = ""
)

$venvPython = "$PSScriptRoot\.venv\Scripts\python.exe"
if (Test-Path $venvPython) {
    $pythonCmd = $venvPython
} else {
    $pythonCmd = "python"
}

if ($Version -and $Message) {
    & $pythonCmd "$PSScriptRoot\release.py" $Version -m $Message
} elseif ($Version) {
    & $pythonCmd "$PSScriptRoot\release.py" $Version
} else {
    & $pythonCmd "$PSScriptRoot\release.py"
}
