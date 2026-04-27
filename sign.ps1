# Скрипт для подписи .exe сертификатом "Savitskiy Artem".
# Запускать после каждой пересборки: .\sign.ps1
# При пересборке через dotnet publish старая подпись теряется,
# поэтому нужно подписывать заново.

$ErrorActionPreference = "Stop"

$certSubject = "CN=Savitskiy Artem"
$exePath = Join-Path $PSScriptRoot "DocxToPdfConverter.exe"
$timestampServer = "http://timestamp.digicert.com"

if (-not (Test-Path $exePath)) {
    Write-Error "Не найден $exePath. Сначала выполните: dotnet publish -c Release"
    exit 1
}

$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $certSubject } |
    Select-Object -First 1

if (-not $cert) {
    Write-Error "Сертификат '$certSubject' не найден в хранилище CurrentUser\My."
    exit 1
}

$result = Set-AuthenticodeSignature `
    -FilePath $exePath `
    -Certificate $cert `
    -TimestampServer $timestampServer `
    -HashAlgorithm SHA256

Write-Host "Подписан: $exePath"
Write-Host "Издатель: $($result.SignerCertificate.Subject)"
Write-Host "Status:   $($result.Status)"
if ($result.Status -ne "Valid") {
    Write-Host "(Status 'UnknownError' — это нормально для самоподписанного сертификата." -ForegroundColor Yellow
    Write-Host " Сама подпись на файле есть. Чтобы Windows показывал её как доверенную," -ForegroundColor Yellow
    Write-Host " установите Cert\SavitskiyArtem.cer в 'Доверенные корневые'.)" -ForegroundColor Yellow
}
