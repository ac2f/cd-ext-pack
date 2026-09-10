<#
.SYNOPSIS
    ac2f pack kaynak dosyalarini CorelDRAW VBA'ya aktarilmaya hazir hale getirir.

.DESCRIPTION
    1.2.0'dan itibaren kaynak saf ASCII'dir, bu yuzden kod sayfasi donusumu
    ZORUNLU DEGILDIR; src\*.bas dosyalari dogrudan ice aktarilabilir.
    Bu betik yalnizca CRLF satir sonlu kopya uretmek icin kolaylik saglar.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\build.ps1
#>
[CmdletBinding()]
param(
    [string]$Source = (Join-Path $PSScriptRoot '..\src'),
    [string]$Output = (Join-Path $PSScriptRoot '..\build'),
    [int]$CodePage = 1254
)

$ErrorActionPreference = 'Stop'

$Source = (Resolve-Path $Source).Path
if (-not (Test-Path $Output)) { New-Item -ItemType Directory -Path $Output | Out-Null }
$Output = (Resolve-Path $Output).Path

$enc = [System.Text.Encoding]::GetEncoding($CodePage)

Get-ChildItem -Path $Source -Filter '*.bas' | ForEach-Object {
    $text = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
    $text = $text -replace "`r`n", "`n" -replace "`n", "`r`n"
    $dest = Join-Path $Output $_.Name
    [System.IO.File]::WriteAllText($dest, $text, $enc)
    Write-Host ("  {0,-20} -> {1}" -f $_.Name, $dest)
}

Write-Host ''
Write-Host "Hazir. CorelDRAW > Araclar > Makrolar > Makro Duzenleyici (Alt+F11)"
Write-Host "icinde ac2fPack projesini secip File > Import File ile build/ altindaki"
Write-Host ".bas dosyalarini iceri aktarin."
