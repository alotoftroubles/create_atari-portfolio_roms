# Check if running as admin
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Script must be run as Administrator. Right-click and select 'Run as administrator'." -ForegroundColor Red
    exit 1
}

# Process ZIP files in current directory
$zipFiles = Get-ChildItem -Filter "*.zip"
if ($zipFiles.Count -eq 0) { Write-Host "No ZIP files found."; exit }

foreach ($zip in $zipFiles) {
    $romName = [System.IO.Path]::GetFileNameWithoutExtension($zip.Name) + ".rom"
    $romPath = Join-Path (Get-Location) $romName
    
    # Copy template ROM
    Copy-Item -Force "128kb-fat12.rom" $romPath
    
    # Mount ROM as virtual disk (ImDisk required: https://sourceforge.net/projects/imdisk-toolkit/)
    $mountLetter = "R:"
    & imdisk -a -f $romPath -m $mountLetter -p "/fs:fat12 /q /y" 2>$null
    Start-Sleep -Seconds 2
    
    # Extract ZIP to mounted drive
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zip.FullName, $mountLetter + "\")
    
    # Create/update autoexec.bat
    "@echo off`n" | Out-File -FilePath "$mountLetter\autoexec.bat" -Encoding ASCII
    
    # Find executables (*.com, *.exe case-insensitive)
    $executables = Get-ChildItem -Path $mountLetter -Include *.com, *.exe -Recurse | Where-Object { $_.Extension -match 'com|exe' }
    Write-Host "`nExecutables found:"
    $executables.Name | ForEach-Object { Write-Host $_ }
    
    if ($executables.Count -eq 1) {
        $exeName = [System.IO.Path]::GetFileNameWithoutExtension($executables[0].Name)
        Write-Host "`nOne executable detected, adding to autoexec.bat: $exeName"
        "`n$exeName" | Add-Content -Path "$mountLetter\autoexec.bat" -Encoding ASCII
    } else {
        Write-Host "`nEnter command(s) for autoexec.bat (multi-line OK, e.g., 'cd subdir`nprogram.exe'):"
        $input = Read-Host
        $input | Add-Content -Path "$mountLetter\autoexec.bat" -Encoding ASCII
    }
    
    # Unmount
    & imdisk -D -m $mountLetter 2>$null
    Start-Sleep -Seconds 2
    
    Write-Host "Created $romName`n"
}

Write-Host "Done."
