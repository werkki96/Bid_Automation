# 파일 역할: Python이 없는 PC에서도 실행되는 포터블 배포 폴더와 zip을 만듭니다.
# 마지막 업데이트: 2026-11-03
# 작업자: SMI 장지수
#
# PyInstaller EXE는 서명이 없어 스마트 앱 컨트롤(SAC)이 차단합니다.
# 이 배포본은 python.org가 서명한 pythonw.exe로 앱 소스를 실행하므로 SAC·SmartScreen을 통과합니다.

$ErrorActionPreference = "Stop"
$PyVersion = "3.14.10"
$PyTag = "python" + ($PyVersion.Split(".")[0..1] -join "")   # python314

$Root = $PSScriptRoot
$Cache = Join-Path $Root "build\portable_cache\$PyVersion"
$Out = Join-Path $Root "dist\BidNoticeCollector_portable"
$Runtime = Join-Path $Out "runtime"
$App = Join-Path $Out "app"
$Zip = Join-Path $Root "dist\BidNoticeCollector_portable.zip"

# 1) python.org 임베디드 Python과 Tcl/Tk(MSI)를 내려받습니다. 한 번 받은 파일은 재사용합니다.
New-Item -ItemType Directory -Force $Cache | Out-Null
$EmbedZip = Join-Path $Cache "embed.zip"
$TclMsi = Join-Path $Cache "tcltk.msi"
$TclDir = Join-Path $Cache "tcltk"
if (-not (Test-Path $EmbedZip)) {
    Invoke-WebRequest "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-embed-amd64.zip" -OutFile $EmbedZip
}
if (-not (Test-Path $TclMsi)) {
    Invoke-WebRequest "https://www.python.org/ftp/python/$PyVersion/amd64/tcltk.msi" -OutFile $TclMsi
}
if (-not (Test-Path $TclDir)) {
    # 관리 설치(/a)는 PC에 설치·등록하지 않고 파일만 풀어냅니다.
    $p = Start-Process msiexec -ArgumentList "/a `"$TclMsi`" /qn TARGETDIR=`"$TclDir`"" -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw "tcltk.msi 추출 실패 (exit $($p.ExitCode))" }
}

# 2) 런타임을 조립합니다. 임베디드 패키지에는 tkinter가 없어 Tcl/Tk 파일을 더합니다.
if (Test-Path $Out) { Remove-Item $Out -Recurse -Force }
Expand-Archive $EmbedZip -DestinationPath $Runtime
Copy-Item (Join-Path $TclDir "DLLs\*") $Runtime -Force
Copy-Item (Join-Path $TclDir "tcl") $Runtime -Recurse
New-Item -ItemType Directory -Force (Join-Path $Runtime "Lib") | Out-Null
Copy-Item (Join-Path $TclDir "Lib\tkinter") (Join-Path $Runtime "Lib") -Recurse
Set-Content (Join-Path $Runtime "$PyTag._pth") -Encoding ascii -Value @(
    "$PyTag.zip", ".", "Lib", "Lib\site-packages", "..\app"
)

# 3) openpyxl 등 순수 Python 의존성을 런타임 안에 설치합니다.
python -m pip install --quiet --no-compile --only-binary=:all: --platform win_amd64 `
    --python-version $PyVersion --target (Join-Path $Runtime "Lib\site-packages") -r (Join-Path $Root "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "의존성 설치 실패" }

# 4) 앱 소스와 사용자 설정을 복사합니다. 설정·인증키·결과는 폴더 최상위에 둡니다.
New-Item -ItemType Directory -Force $App | Out-Null
Copy-Item (Join-Path $Root "*.py") $App
Copy-Item (Join-Path $Root "collectors") $App -Recurse
Get-ChildItem $App -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force
Copy-Item (Join-Path $Root "settings.json") $Out
foreach ($name in ".env", ".env.example") {
    $src = Join-Path $Root $name
    if (Test-Path $src) { Copy-Item $src $Out }
}
New-Item -ItemType Directory -Force (Join-Path $Out "output") | Out-Null

# 5) 실행용 .bat을 만듭니다. cmd 코드페이지 문제를 피하려고 내용은 ASCII로만 씁니다.
Set-Content (Join-Path $Out "BidNoticeCollector 실행.bat") -Encoding ascii -Value @(
    "@echo off",
    "set `"BID_NOTICE_HOME=%~dp0.`"",
    "start `"`" `"%~dp0runtime\pythonw.exe`" `"%~dp0app\desktop_app.py`""
)

# 6) 배포용 zip을 만듭니다. 방금 푼 파일을 백신이 검사 중이면 잠길 수 있어 몇 번 재시도합니다.
Add-Type -AssemblyName System.IO.Compression.FileSystem
for ($try = 1; ; $try++) {
    if (Test-Path $Zip) { Remove-Item $Zip -Force }
    try {
        [IO.Compression.ZipFile]::CreateFromDirectory($Out, $Zip, [IO.Compression.CompressionLevel]::Optimal, $true)
        break
    } catch {
        if ($try -ge 5) { throw }
        Start-Sleep -Seconds 3
    }
}
Write-Host "Portable folder: $Out"
Write-Host "Portable zip:    $Zip"
