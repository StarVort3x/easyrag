Param(
    [string]$CondaEnv = "easyrag",
    [string]$BackendDir = "backend",
    [string]$FrontendDir = "ruoyi-ui"
)

$ErrorActionPreference = "Stop"

function Resolve-PathSafe {
    param([string]$PathSegment)
    return (Join-Path -Path $PSScriptRoot -ChildPath $PathSegment)
}

$backendPath = Resolve-PathSafe $BackendDir
$frontendPath = Resolve-PathSafe $FrontendDir

if (-not (Test-Path $backendPath)) {
    Write-Error "后端目录不存在：$backendPath"
}

if (-not (Test-Path $frontendPath)) {
    Write-Error "前端目录不存在：$frontendPath"
}

Write-Host "后端目录: $backendPath"
Write-Host "前端目录: $frontendPath"
Write-Host ""

$activateConda = @"
if (Get-Command conda -ErrorAction SilentlyContinue) {
    conda activate $CondaEnv
} elseif (Get-Command "$Env:USERPROFILE\anaconda3\Scripts\activate.bat" -ErrorAction SilentlyContinue) {
    & "$Env:USERPROFILE\anaconda3\Scripts\activate.bat" $CondaEnv
} else {
    Write-Warning '未找到 conda，默认使用当前环境'
}
"@

$backendCmd = @"
$activateConda
cd "$backendPath"
python start.py
"@

$frontendCmd = @"
cd "$frontendPath"
npm run serve
"@

Write-Host "启动 FastAPI 后端..."
Start-Process powershell -ArgumentList "-NoExit","-Command",$backendCmd

Write-Host "启动 Vue 前端..."
Start-Process powershell -ArgumentList "-NoExit","-Command",$frontendCmd

Write-Host ""
Write-Host "所有进程已启动。关闭各自窗口即可停止服务。"

