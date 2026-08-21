param(
  [string]$Source = 'D:\桌面\WK34_原工具界面版_1比1.html',
  [string]$CompanyRoot = 'D:\桌面\GIT\公司\Company-'
)

$ErrorActionPreference = 'Stop'
$target = Join-Path $CompanyRoot 'deploy\shop\index.html'

if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
  throw "店铺分析 HTML 不存在：$Source"
}

New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
Copy-Item -LiteralPath $Source -Destination $target -Force

$sourceInfo = Get-Item -LiteralPath $Source
$targetInfo = Get-Item -LiteralPath $target
if ($sourceInfo.Length -ne $targetInfo.Length) {
  throw "同步校验失败：源文件和公司入口文件大小不一致。"
}

Write-Output "已同步店铺分析：$target"
Write-Output "文件大小：$($targetInfo.Length) bytes"
