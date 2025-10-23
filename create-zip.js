const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

// Use PowerShell to create zip with proper structure
const psScript = `
$source = Join-Path $PSScriptRoot "dist"
$destination = Join-Path $PSScriptRoot "d365-helper-extension.zip"

# Remove existing zip
if (Test-Path $destination) {
    Remove-Item $destination -Force
}

# Create zip with proper paths (no backslashes issue)
Add-Type -Assembly System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($source, $destination, [System.IO.Compression.CompressionLevel]::Optimal, $false)

Write-Host "Created: $destination"
Get-Item $destination | Select-Object Name, Length
`;

// Write temp PS script
fs.writeFileSync('temp-zip.ps1', psScript);

// Execute it
exec('powershell -ExecutionPolicy Bypass -File temp-zip.ps1', (error, stdout, stderr) => {
    // Clean up temp file
    try { fs.unlinkSync('temp-zip.ps1'); } catch(e) {}

    if (error) {
        console.error('Error creating zip:', error);
        process.exit(1);
    }

    console.log(stdout);
    if (stderr) console.error(stderr);

    console.log('\n✓ Extension packaged successfully!');
    console.log('File: d365-helper-extension.zip');
});
