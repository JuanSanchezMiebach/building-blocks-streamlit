# ============================================================
# EJECUTAR APP STREAMLIT - GENERADOR DE BUILDING BLOCKS V6
# ============================================================

$ErrorActionPreference = "Stop"
$AppPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $AppPath

Write-Host "" 
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " EJECUTANDO GENERADOR DE BUILDING BLOCKS V6" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Carpeta: $AppPath"
Write-Host ""

if (!(Test-Path "app.py")) {
    Write-Host "ERROR: No se encontro app.py en esta carpeta." -ForegroundColor Red
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

if (!(Test-Path "bb_core.py")) {
    Write-Host "ERROR: No se encontro bb_core.py en esta carpeta." -ForegroundColor Red
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

if (!(Test-Path "requirements.txt")) {
@"
streamlit>=1.35
pandas>=2.0
numpy>=1.24
openpyxl>=3.1
xlsxwriter>=3.2
altair>=5.0
"@ | Out-File -FilePath "requirements.txt" -Encoding utf8
}

# Liberar puerto 8501 si quedo una app anterior abierta
try {
    $connections = Get-NetTCPConnection -LocalPort 8501 -ErrorAction SilentlyContinue
    foreach ($conn in $connections) {
        if ($conn.OwningProcess -and $conn.OwningProcess -ne $PID) {
            Write-Host "Cerrando proceso anterior en puerto 8501: $($conn.OwningProcess)" -ForegroundColor Yellow
            Stop-Process -Id $conn.OwningProcess -Force -ErrorAction SilentlyContinue
        }
    }
} catch {
    Write-Host "No fue posible validar el puerto 8501. Continuando..." -ForegroundColor Yellow
}

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$PythonCmd = $null
try {
    py --version | Out-Null
    $PythonCmd = "py"
} catch {
    try {
        python --version | Out-Null
        $PythonCmd = "python"
    } catch {
        Write-Host "ERROR: No se encontro Python. Instala Python y marca 'Add Python to PATH'." -ForegroundColor Red
        Read-Host "Presiona Enter para cerrar"
        exit 1
    }
}

if (!(Test-Path ".venv")) {
    Write-Host "Creando ambiente virtual .venv..." -ForegroundColor Cyan
    & $PythonCmd -m venv .venv
}

if (!(Test-Path ".\.venv\Scripts\Activate.ps1")) {
    Write-Host "ERROR: No se encontro .venv\Scripts\Activate.ps1" -ForegroundColor Red
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

. .\.venv\Scripts\Activate.ps1

Write-Host "Actualizando pip..." -ForegroundColor Cyan
python -m pip install --upgrade pip

Write-Host "Instalando dependencias..." -ForegroundColor Cyan
python -m pip install -r requirements.txt

Write-Host "Validando sintaxis..." -ForegroundColor Cyan
python -m py_compile app.py bb_core.py

if (!(Test-Path "data")) {
    New-Item -ItemType Directory -Path "data" | Out-Null
}

Write-Host "" 
Write-Host "============================================================" -ForegroundColor Green
Write-Host " APP LISTA - http://localhost:8501" -ForegroundColor Green
Write-Host " Para detenerla usa CTRL + C" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "" 

python -m streamlit run app.py --server.port 8501
