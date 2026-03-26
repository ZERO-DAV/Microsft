# -- Step 1: Require Administrator Privileges --
if (-NOT ([Security.Principal.WindowsPrincipal]
         [Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {

    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

# -- Step 2: Create C:\Microsoft if missing --
$dirPath = "C:\Microsoft"

if (-not (Test-Path $dirPath)) {
    New-Item -Path $dirPath -ItemType Directory | Out-Null
    Write-Host "Folder created: C:\Microsoft" -ForegroundColor Green
}

# -- Step 3: Download Files from GitHub --
$baseUrl = "https://raw.githubusercontent.com/ZERO-DAV/Microsft/main"
$files   = @(
    "Configuration.xml",
    "configuration-Office365-x64.xml",
    "officedeploymenttool_19725-20126.exe",
    "setup.exe"
)

Write-Host "Downloading files from ZERO-DAV repository..." -ForegroundColor Cyan

foreach ($file in $files) {
    $url  = "$baseUrl/$file"
    $dest = Join-Path $dirPath $file
    Write-Host "  Downloading: $file" -ForegroundColor Gray
    Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
}

Write-Host "All files downloaded successfully." -ForegroundColor Green

# -- Step 4: Open CMD as Admin -> cd into folder -> run setup --
$setupPath  = Join-Path $dirPath "setup.exe"
$configPath = Join-Path $dirPath "configuration.xml"

if (Test-Path $setupPath) {
    Write-Host "Opening CMD and running installer..." -ForegroundColor Yellow

    $cmd = "cd /d C:\Microsoft && setup.exe /configure configuration.xml && pause"

    Start-Process "cmd.exe" `
        -ArgumentList "/K $cmd" `
        -Verb RunAs `
        -WorkingDirectory $dirPath
} else {
    Write-Host "Error: setup.exe not found in C:\Microsoft" -ForegroundColor Red
}
