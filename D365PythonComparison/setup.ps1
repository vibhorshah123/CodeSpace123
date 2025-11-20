# Setup script for D365 Python Comparison Tool
# Run this script to set up your environment

Write-Host "=" -NoNewline -ForegroundColor Cyan
Write-Host ("=" * 68) -ForegroundColor Cyan
Write-Host "  D365 Python Comparison Tool - Setup" -ForegroundColor White
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""

# Check Python installation
Write-Host "Checking Python installation..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "  ✓ Found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Python not found! Please install Python 3.8 or higher." -ForegroundColor Red
    Write-Host "    Download from: https://www.python.org/downloads/" -ForegroundColor Yellow
    exit 1
}

# Check if pip is available
Write-Host "Checking pip installation..." -ForegroundColor Yellow
try {
    $pipVersion = pip --version 2>&1
    Write-Host "  ✓ Found: $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "  ✗ pip not found!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Installing required Python packages..." -ForegroundColor Yellow
Write-Host "This may take a few moments..." -ForegroundColor Gray
Write-Host ""

# Install requirements
try {
    pip install -r requirements.txt
    Write-Host ""
    Write-Host "  ✓ All packages installed successfully!" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Failed to install packages!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host "  Setup Complete!" -ForegroundColor Green
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""
Write-Host "To run the tool, execute:" -ForegroundColor White
Write-Host "  python main.py" -ForegroundColor Cyan
Write-Host ""
Write-Host "For more information, see README.md" -ForegroundColor Gray
Write-Host ""
