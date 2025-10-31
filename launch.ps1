# Microsoft Graph Security Webhook Tester - Launch Script
# This script sets up the environment and launches the application

Write-Host "Microsoft Graph Security Webhook Tester" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Python not found. Please install Python 3.8+ from https://python.org" -ForegroundColor Red
    pause
    exit 1
}

# Get the directory where this script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

Write-Host "Working directory: $scriptDir" -ForegroundColor Yellow

# Check if virtual environment exists
if (Test-Path ".venv") {
    Write-Host "Virtual environment found" -ForegroundColor Green
} else {
    Write-Host "Virtual environment not found. Creating one..." -ForegroundColor Yellow
    python -m venv .venv
    Write-Host "Virtual environment created" -ForegroundColor Green
}

# Activate virtual environment
Write-Host "Activating virtual environment..." -ForegroundColor Yellow
& ".venv\Scripts\Activate.ps1"

# Install/upgrade requirements
Write-Host "Installing/updating requirements..." -ForegroundColor Yellow
pip install -r requirements.txt

# Check if config file exists
if (Test-Path "config.json") {
    Write-Host "Configuration file found" -ForegroundColor Green
} else {
    Write-Host "Configuration file not found. Using template..." -ForegroundColor Yellow
    Copy-Item "config_template.json" "config.json"
    Write-Host "Please edit config.json with your Azure app registration details" -ForegroundColor Cyan
    Write-Host "You can also configure it through the application GUI" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "Launching Graph Security Webhook Tester..." -ForegroundColor Green
Write-Host ""

# Launch the application
python graph_security_webhook_tester.py

Write-Host ""
Write-Host "Application closed" -ForegroundColor Yellow
pause