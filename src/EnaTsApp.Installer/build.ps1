param(
    [string]$Configuration = "Release"
)

# Build the main application
Write-Host "Building EnaTsApp..."
& dotnet build ..\EnaTsApp\EnaTsApp.csproj -c $Configuration

# Create installer directory
$installerDir = Join-Path $PSScriptRoot "InstallerBuild"
if (Test-Path $installerDir) {
    Remove-Item $installerDir -Recurse -Force
}
New-Item -ItemType Directory -Path $installerDir

# Copy application files
Write-Host "Copying application files..."
Copy-Item -Path "..\EnaTsApp\bin\$Configuration\net7.0\*" -Destination $installerDir -Recurse

# Generate WiX installer
Write-Host "Generating WiX installer..."
& heat.exe dir $installerDir -nologo -gg -sfrag -sreg -srd -scom -dr INSTALLFOLDER -var var.SourceDir -out Components.wxs

# Build the installer
Write-Host "Building installer..."
& candle.exe Product.wxs Components.wxs
& light.exe -ext WixUIExtension -out EnaTsAppSetup.msi Product.wixobj Components.wixobj

Write-Host "Installer created successfully!" -ForegroundColor Green
