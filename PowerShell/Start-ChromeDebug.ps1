param(
    [string]$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe",
    [string]$ProfileDir = "C:\Users\necul\AppData\Local\Google\Chrome\User Data\Default",
    [int]$DebugPort = 9222,
    [string]$Url = "https://translate.google.ro/?hl=ro&sl=auto&tl=ro&op=docs",
    [switch]$SkipCleanup
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ChromePath)) {
    throw "Chrome nu exista la: $ChromePath"
}

if (-not $SkipCleanup) {
    Write-Host "[CLEANUP] Opresc Chrome si ChromeDriver..."
    Get-Process -Name chrome, chromedriver -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

Write-Host "[START] Chrome debug port $DebugPort"
Write-Host "[START] Profil: $ProfileDir"

$chromeArgs = @(
    "--remote-debugging-port=$DebugPort",
    "--user-data-dir=$ProfileDir",
    $Url
)

Start-Process -FilePath $ChromePath -ArgumentList $chromeArgs -WindowStyle Normal

$versionUrl = "http://127.0.0.1:$DebugPort/json/version"
for ($i = 1; $i -le 60; $i++) {
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $versionUrl -TimeoutSec 2
        if ($response.StatusCode -eq 200) {
            Write-Host "[OK] Chrome debug raspunde pe $versionUrl"
            exit 0
        }
    }
    catch {
        Start-Sleep -Seconds 1
    }
}

throw "Chrome debug nu a raspuns pe $versionUrl in 60 secunde"
