#requires -RunAsAdministrator
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Change these defaults if needed.
$language = 'en-us'
$channel = 'MonthlyEnterprise'
$product = 'O365ProPlusRetail'
$root = Join-Path $env:ProgramData ('AVD-Office32-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
Start-Transcript -Path (Join-Path $root 'Install.log') -Force | Out-Null

try {
    if (-not [Environment]::Is64BitProcess) {
        throw 'Run this script in 64-bit PowerShell. It installs 32-bit Office.'
    }

    $regPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $existing = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    if ($existing.Platform -eq 'x64') {
        throw '64-bit Click-to-Run Office is already installed. Use a base image without Office, or explicitly plan an Office architecture migration first.'
    }

    # Resolve the current ODT installer from Microsoft's official download page.
    $downloadPage = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' -UseBasicParsing
    $html = [System.Net.WebUtility]::HtmlDecode($downloadPage.Content)
    $pattern = 'https://download\.microsoft\.com/[^"\s<>]+/officedeploymenttool[^/"\s<>]*\.exe'
    $downloadUrl = [regex]::Match($html, $pattern, 'IgnoreCase').Value
    if (-not $downloadUrl) {
        throw 'Could not resolve the Office Deployment Tool download URL. Microsoft may have changed the download page.'
    }

    $odt = Join-Path $root 'officedeploymenttool.exe'
    Invoke-WebRequest -Uri $downloadUrl -OutFile $odt -UseBasicParsing
    $signature = Get-AuthenticodeSignature -FilePath $odt
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        throw 'The downloaded Office Deployment Tool does not have a valid Microsoft signature.'
    }

    $extract = Start-Process -FilePath $odt -ArgumentList "/quiet /extract:`"$root`"" -Wait -PassThru -WindowStyle Hidden
    if ($extract.ExitCode -ne 0) { throw "ODT extraction failed: $($extract.ExitCode)" }
    $setup = Join-Path $root 'setup.exe'
    if (-not (Test-Path -LiteralPath $setup)) { throw 'ODT setup.exe was not extracted.' }

    $configPath = Join-Path $root 'Office32.xml'
    @"
<Configuration>
  <Add OfficeClientEdition="32" Channel="$channel">
    <Product ID="$product">
      <Language ID="$language" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Teams" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@ | Set-Content -LiteralPath $configPath -Encoding UTF8

    # Configure downloads directly from the Office CDN and installs silently.
    $install = Start-Process -FilePath $setup -ArgumentList "/configure `"$configPath`"" -Wait -PassThru -WindowStyle Hidden
    if ($install.ExitCode -notin @(0, 3010)) {
        throw "Office installation failed with exit code $($install.ExitCode). See Office setup logs in $env:TEMP and transcript in $root."
    }

    $installed = Get-ItemProperty -Path $regPath
    if ($installed.Platform -ne 'x86') { throw 'Verification failed: Office platform is not x86.' }
    if ([string]$installed.SharedComputerLicensing -ne '1') { throw 'Verification failed: shared computer activation is not enabled.' }
    if (($installed.ProductReleaseIds -split ',') -notcontains $product) { throw "Verification failed: $product is not installed." }
    $wordPath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\WINWORD.EXE'
    if (-not (Test-Path -LiteralPath $wordPath)) { throw 'Verification failed: 32-bit Word executable was not found.' }

    Write-Output "Verified: 32-bit Microsoft 365 Apps with shared computer activation. Version: $($installed.VersionToReport)"
    Write-Output "Installation records: $root"
    if ($install.ExitCode -eq 3010) { Write-Output 'Office requested a restart. Add a WindowsRestart customizer after this step.' }
}
catch {
    Write-Output "FAILED: $($_.Exception.Message)"
    throw
}
finally {
    Stop-Transcript | Out-Null
}
