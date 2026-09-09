# Markdown 工具箱 — Word 加载项安装脚本
# ------------------------------------------------------------------
# 用法（在 PowerShell 中以当前用户身份运行，无需管理员）：
#   .\install.ps1
#
# 脚本会：
#   1. 检查 Node.js
#   2. 生成自签名 HTTPS 证书（若不存在）
#   3. 把本加载项注册到 Word 的受信任目录
#
# 卸载：.\install.ps1 -Uninstall

param(
  [switch]$Uninstall,
  [int]$Port = 3000
)

$ErrorActionPreference = 'Stop'
$AddinDir = $PSScriptRoot
$ManifestPath = Join-Path $AddinDir 'manifest.xml'
$CertDir = Join-Path $AddinDir '.certs'

function Write-Step($msg) { Write-Host "  $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Write-Warn2($msg) { Write-Host "  [!] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "  [X] $msg" -ForegroundColor Red }

# ------------------------------------------------------------------
# 卸载
# ------------------------------------------------------------------
$AddinGuid = 'c572bc3f-65ba-4b89-9360-f68811b63a3b'
$CatalogGuid = '{7C4E9B21-3F6A-4D8E-9A15-2B7C0D5E8F31}'

if ($Uninstall) {
  Write-Host "`n正在卸载 Markdown 工具箱...`n" -ForegroundColor White
  $catalogPath = Join-Path $env:LOCALAPPDATA 'MarkdownToolbox\catalog'
  if (Test-Path $catalogPath) {
    Remove-Item $catalogPath -Recurse -Force
    Write-Ok "已移除本地目录注册：$catalogPath"
  }
  foreach ($ver in @('16.0', '15.0')) {
    $wef = "HKCU:\Software\Microsoft\Office\$ver\WEF"
    Remove-ItemProperty -Path "$wef\Developer" -Name $AddinGuid -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "$wef\Developer" -Name 'TrustedCatalogs' -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "$wef\Developer" -Name 'RefreshAddins' -ErrorAction SilentlyContinue
    Remove-Item -Path "$wef\TrustedCatalogs\$CatalogGuid" -Recurse -Force -ErrorAction SilentlyContinue
  }
  Write-Ok "已清理注册表项"
  $wefCache = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
  if (Test-Path $wefCache) {
    Get-ChildItem $wefCache -Directory -ErrorAction SilentlyContinue | ForEach-Object {
      $dev = Join-Path $_.FullName 'DeveloperSettings'
      if (Test-Path $dev) {
        Remove-Item $dev -Recurse -Force -ErrorAction SilentlyContinue
        Write-Ok "已清理缓存：$dev"
      }
    }
  }
  Write-Host "`n卸载完成。请重启 Word。`n" -ForegroundColor Green
  exit 0
}

# ------------------------------------------------------------------
# 安装
# ------------------------------------------------------------------
Write-Host "`n===== Markdown 工具箱 · Word 加载项安装 =====`n" -ForegroundColor White

# 1. 检查清单
Write-Step "检查清单文件..."
if (-not (Test-Path $ManifestPath)) {
  Write-Err "找不到 manifest.xml：$ManifestPath"
  exit 1
}
Write-Ok "manifest.xml 存在"

# 2. 检查 Node
Write-Step "检查 Node.js..."
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
  Write-Warn2 "未检测到 Node.js。你仍可使用加载项，但需要其他方式启动 HTTPS 服务。"
  Write-Warn2 "推荐安装 Node.js：https://nodejs.org/"
} else {
  Write-Ok "Node.js $(& node --version)"
}

# 3. 生成证书
Write-Step "准备 HTTPS 证书..."
$keyPath = Join-Path $CertDir 'key.pem'
$certPath = Join-Path $CertDir 'cert.pem'
if ((Test-Path $keyPath) -and (Test-Path $certPath)) {
  Write-Ok "证书已存在"
} else {
  $openssl = Get-Command openssl -ErrorAction SilentlyContinue
  if ($openssl) {
    New-Item -ItemType Directory -Force -Path $CertDir | Out-Null
    $confPath = Join-Path $CertDir 'openssl.cnf'
    @'
[req]
distinguished_name = dn
x509_extensions = v3
prompt = no
[dn]
CN = localhost
[v3]
subjectAltName = DNS:localhost,IP:127.0.0.1
basicConstraints = CA:FALSE
keyUsage = digitalSignature,keyEncipherment
extendedKeyUsage = serverAuth
'@ | Set-Content -Path $confPath -Encoding ASCII

    & openssl req -x509 -newkey rsa:2048 -nodes `
        -keyout $keyPath -out $certPath -days 825 -config $confPath 2>$null
    if (Test-Path $certPath) {
      Write-Ok "已生成自签名证书"
      Write-Step "正在信任证书（浏览器将不再提示）..."
      $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
      $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
      $store.Open('ReadWrite')
      $store.Add($cert)
      $store.Close()
      Write-Ok "证书已加入当前用户受信任根证书"
    } else {
      Write-Warn2 "证书生成失败，服务器启动时会重试"
    }
  } else {
    Write-Warn2 "未找到 openssl，服务器首次启动时会尝试生成证书"
  }
}

# 4. 创建本地目录并注册（双通道）
Write-Step "注册加载项到 Word..."
$shareDir = Join-Path $env:LOCALAPPDATA 'MarkdownToolbox\catalog'
New-Item -ItemType Directory -Force -Path $shareDir | Out-Null
Copy-Item $ManifestPath (Join-Path $shareDir 'manifest.xml') -Force
$manifestFullPath = Join-Path $shareDir 'manifest.xml'
Write-Ok "清单已复制到：$manifestFullPath"

$okDev = $false
$okCatalog = $false

# 找一个能覆盖 $shareDir 的既有共享，构造 UNC 路径（共享文件夹通道要求 UNC）
$uncUrl = $null
$shares = @()
try {
  $shareDefs = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares' -ErrorAction Stop
  foreach ($p in $shareDefs.PSObject.Properties) {
    if ($p.Name -match '^PS' -or $p.Name -eq 'WefCacheId') { continue }
    $raw = ($p.Value -join ';')
    $mPath = [regex]::Match($raw, 'Path=([^;\0]+)')
    $mName = [regex]::Match($raw, 'ShareName=([^;\0]+)')
    if ($mPath.Success -and $mName.Success) {
      $shares += [pscustomobject]@{ Name = $mName.Groups[1].Value; Path = $mPath.Groups[1].Value }
    }
  }
} catch {
  try {
    Get-SmbShare -ErrorAction Stop | Where-Object { -not $_.Name.EndsWith('$') } |
      ForEach-Object { $shares += [pscustomobject]@{ Name = $_.Name; Path = $_.Path } }
  } catch { }
}

$hostName = $env:COMPUTERNAME
if (-not $hostName) { $hostName = [System.Net.Dns]::GetHostName() }
$bestLen = -1
foreach ($s in $shares) {
  $base = $s.Path.TrimEnd('\')
  if ($shareDir.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) {
    if ($base.Length -gt $bestLen) {
      $bestLen = $base.Length
      $rel = $shareDir.Substring($base.Length).TrimStart('\')
      if ($rel) {
        $uncUrl = "\\$hostName\$($s.Name)\$rel"
      } else {
        $uncUrl = "\\$hostName\$($s.Name)"
      }
    }
  }
}

foreach ($ver in @('16.0', '15.0')) {
  $wef = "HKCU:\Software\Microsoft\Office\$ver\WEF"
  $devKey = "$wef\Developer"
  if (-not (Test-Path $devKey)) { New-Item -Path $devKey -Force | Out-Null }

  # 通道 B：开发者直挂清单（微软官方 office-addin-dev-settings 用的就是这条）
  Set-ItemProperty -Path $devKey -Name $AddinGuid -Value $manifestFullPath -Type String
  Set-ItemProperty -Path $devKey -Name 'RefreshAddins' -Value 1 -Type DWord
  # 清理旧脚本写错的残留值
  Remove-ItemProperty -Path $devKey -Name 'TrustedCatalogs' -ErrorAction SilentlyContinue
  $okDev = $true

  # 通道 A：共享文件夹目录（Url 必须是 UNC 网络路径）
  if ($uncUrl) {
    $catKey = "$wef\TrustedCatalogs\$CatalogGuid"
    New-Item -Path $catKey -Force | Out-Null
    Set-ItemProperty -Path $catKey -Name 'Id'    -Value $CatalogGuid -Type String
    Set-ItemProperty -Path $catKey -Name 'Url'   -Value $uncUrl      -Type String
    Set-ItemProperty -Path $catKey -Name 'Flags' -Value 1            -Type DWord
    $okCatalog = $true
  }
}

if ($okDev) { Write-Ok "已挂载到 Word（开发者通道）" }
if ($okCatalog) { Write-Ok "已登记共享文件夹目录：$uncUrl" }
if (-not $okDev -and -not $okCatalog) { Write-Warn2 "注册表写入失败，请在 Word 中手动加载清单" }

# 5. 完成
Write-Host "`n===== 安装完成 =====`n" -ForegroundColor Green
Write-Host "接下来：" -ForegroundColor White
Write-Host "  1. 启动开发服务器：" -NoNewline; Write-Host " node server.js" -ForegroundColor Yellow
Write-Host "  2. 完全关闭 Word（所有窗口），重新打开一个文档" -ForegroundColor White
Write-Host "  3. 「开始」选项卡最右侧应出现「Markdown 工具箱」组" -ForegroundColor White
Write-Host "  4. 若未出现：「插入」→「我的加载项」→「共享文件夹」→「Markdown 工具箱」→「添加」" -ForegroundColor White
Write-Host ""
