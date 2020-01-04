Set-StrictMode -Version 'Latest'

Get-ChildItem "$PSScriptRoot/public/*.ps1" | ForEach-Object { . $_ }
