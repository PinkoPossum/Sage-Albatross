# PowerShell Script to Setup Python Environment for Cisco Audit

#Check if Python is installed
Write-Host "Checking for Python installation..."
$pythonPath = Get-Command python -ErrorAction SilentlyContinue
if ($null -eq $pythonPath) {
    Write-Host "ERROR: Python is not installed or not in your system's PATH." -ForegroundColor Red
    Write-Host "Please install Python 3 from python.org and ensure it's added to your PATH." -ForegroundColor Yellow
    # Pause to allow the user to read the message before the window closes.
    Read-Host "Press Enter to exit"
    exit
}
Write-Host "Python found at: $($pythonPath.Source)" -ForegroundColor Green

#Define the name for the virtual environment directory
$venvName = "venv"

#Create requirements.txt file for easy dependency management
Write-Host "Creating requirements.txt file..."
try {
    # List of required Python packages
    "openpyxl", "netmiko" | Out-File -FilePath "requirements.txt" -Encoding utf8 -ErrorAction Stop
    Write-Host "requirements.txt created successfully." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Could not create requirements.txt. Please check folder permissions." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

#Create a virtual environment if it doesn't already exist
if (-not (Test-Path -Path $venvName)) {
    Write-Host "Creating Python virtual environment named '$venvName'..."
    try {
        # Use Python's built-in venv module to create the environment
        python -m venv $venvName -ErrorAction Stop
        Write-Host "Virtual environment created successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to create the Python virtual environment." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit
    }
}
else {
    Write-Host "Virtual environment '$venvName' already exists. Skipping creation." -ForegroundColor Yellow
}

#Install required packages using pip from the virtual environment
Write-Host "Installing required Python packages (openpyxl, netmiko) into the virtual environment..."
# Define the path to the pip executable inside the new virtual environment
$pipPath = Join-Path -Path $PSScriptRoot -ChildPath "$venvName\Scripts\pip.exe"

if (Test-Path $pipPath) {
    try {
        # Execute pip from the venv to install packages directly into that environment
        & $pipPath install -r requirements.txt -ErrorAction Stop
        Write-Host "Required packages installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to install Python packages." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit
    }
}
else {
    Write-Host "ERROR: Could not find pip.exe in the virtual environment. Setup failed." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

#Final Instructions for the User
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "Setup is complete!" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "To run the Python script, follow these steps in your PowerShell terminal:"
Write-Host
Write-Host "1. Activate the virtual environment by running:" -ForegroundColor Yellow
Write-Host "   .\$venvName\Scripts\Activate.ps1"
Write-Host "   (Your terminal prompt should change to show '($venvName)')."
Write-Host
Write-Host "2. Run the Python audit script:" -ForegroundColor Yellow
Write-Host "   python cisco_audit.py"
Write-Host
Write-Host "3. When you are finished, deactivate the environment by simply typing:" -ForegroundColor Yellow
Write-Host "   deactivate"
Write-Host
Read-Host "Press Enter to close this setup script."
