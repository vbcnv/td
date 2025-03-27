#Requires -RunAsAdministrator

# Process command-line parameters
Param(
    [Parameter(Mandatory=$false)]
    [switch]$NonInteractive
)

<#
.SYNOPSIS
    VBC NV Workstation Setup Script
.DESCRIPTION
    This script automates the setup of a standard VBC NV workstation by installing
    required software, configuring the desktop environment, and preparing the system.
.NOTES
    Version:        1.1
    Author:         VBC NV IT
    Last Updated:   2025-03-27
#>

# Script configuration
$Config = @{
    WallpaperUrl = "https://raw.githubusercontent.com/vbcnv/td/main/bg-default.png"
    StartBinUrl = "https://raw.githubusercontent.com/vbcnv/td/main/start2.bin"
    InstallFolder = "C:\VBC NV"
    LogFile = "C:\VBC NV\setup_log.txt"
    WallpaperChecksum = "" # Add SHA256 checksum here if known
    StartBinChecksum = "" # Add SHA256 checksum here if known
    
    # Registry settings to apply
    RegistrySettings = @(
        # Disable TaskbarAI
        @{
            Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            Name = "TaskbarAl"
            Value = 0
            Type = "DWord"
        },
        
        # Desktop System Icons Configuration
        # 0 = Show icon, 1 = Hide icon
        
        # Show Computer/This PC icon
        @{
            Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
            Name = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
            Value = 0  # 0 = Show, 1 = Hide
            Type = "DWord"
        },
        
        # Show Control Panel icon
        @{
            Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
            Name = "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"
            Value = 0  # 0 = Show, 1 = Hide
            Type = "DWord"
        },
        
        # Show User's Files icon
        @{
            Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
            Name = "{59031a47-3f72-44a7-89c5-5595fe6b30ee}"
            Value = 0  # 0 = Show, 1 = Hide
            Type = "DWord"
        },
        
        # Show Recycle Bin icon
        @{
            Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
            Name = "{645FF040-5081-101B-9F08-00AA002F954E}"
            Value = 0  # 0 = Show, 1 = Hide
            Type = "DWord"
        },
        
        # Show Network icon
        @{
            Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
            Name = "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"
            Value = 0  # 0 = Show, 1 = Hide
            Type = "DWord"
        }
        
        # Add more registry settings here as needed
        <# Example:
        @{
            Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
            Name = "NoAutoUpdate"
            Value = 1
            Type = "DWord"
        }
        #>
    )
}

# Initialize logging
Function Write-Log {
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    # Create log directory if it doesn't exist
    $logDir = Split-Path -Parent $Config.LogFile
    if (!(Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Output to console with color
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
    }
    
    # Append to log file
    Add-Content -Path $Config.LogFile -Value $logMessage
}

# Check if a command exists
Function Test-CommandExists {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Command
    )
 
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'stop'
 
    Try { 
        if (Get-Command $Command -ErrorAction Stop) { 
            return $true 
        } 
    }
    Catch { 
        return $false 
    }
    Finally { 
        $ErrorActionPreference = $oldPreference 
    }
}

# Set Windows Desktop Wallpaper
Function Set-Wallpaper {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$WallpaperPath
    )

    Write-Log "Setting wallpaper to: $WallpaperPath" -Level "INFO"
    
    Try {
        $code = @' 
        using System.Runtime.InteropServices; 
        namespace Win32 { 
            public class Wallpaper { 
                [DllImport("user32.dll", CharSet=CharSet.Auto)] 
                static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); 
                
                public static void SetWallpaper(string thePath) { 
                    SystemParametersInfo(20, 0, thePath, 3); 
                }
            }
        }
'@

        add-type $code 
        [Win32.Wallpaper]::SetWallpaper($WallpaperPath)
        Write-Log "Wallpaper set successfully" -Level "SUCCESS"
        return $true
    }
    Catch {
        Write-Log "Failed to set wallpaper: $_" -Level "ERROR"
        return $false
    }
}

# Create installation directory
Function Initialize-InstallFolder {
    Param()
    
    Write-Log "Checking for installation folder: $($Config.InstallFolder)" -Level "INFO"
    
    if (!(Test-Path -Path $Config.InstallFolder)) {
        Try {
            Write-Log "Creating installation folder..." -Level "INFO"
            New-Item -Path $Config.InstallFolder -ItemType Directory -Force | Out-Null
            Write-Log "Installation folder created successfully" -Level "SUCCESS"
            return $true
        }
        Catch {
            Write-Log "Failed to create installation folder: $_" -Level "ERROR"
            return $false
        }
    }
    else {
        Write-Log "Installation folder already exists" -Level "INFO"
        return $true
    }
}

# Download file with progress and verification
Function Get-FileWithProgress {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Url,
        
        [Parameter(Mandatory=$true)]
        [string]$OutFile,
        
        [Parameter(Mandatory=$false)]
        [string]$ExpectedChecksum = ""
    )
    
    Write-Log "Downloading: $Url to $OutFile" -Level "INFO"
    
    Try {
        # Create progress bar
        $ProgressPreference = 'Continue'
        
        # Download file
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
        
        # Verify file exists
        if (!(Test-Path -Path $OutFile)) {
            Write-Log "Download failed: File not found at $OutFile" -Level "ERROR"
            return $false
        }
        
        # Verify checksum if provided
        if ($ExpectedChecksum -ne "") {
            $fileHash = Get-FileHash -Path $OutFile -Algorithm SHA256
            if ($fileHash.Hash -ne $ExpectedChecksum) {
                Write-Log "Checksum verification failed!" -Level "ERROR"
                Write-Log "Expected: $ExpectedChecksum" -Level "ERROR"
                Write-Log "Actual: $($fileHash.Hash)" -Level "ERROR"
                return $false
            }
            Write-Log "Checksum verified successfully" -Level "SUCCESS"
        }
        
        Write-Log "Downloaded successfully: $OutFile" -Level "SUCCESS"
        return $true
    }
    Catch {
        Write-Log "Download failed: $_" -Level "ERROR"
        return $false
    }
    Finally {
        $ProgressPreference = 'SilentlyContinue'
    }
}

# Install software packages using winget
Function Install-RequiredSoftware {
    Param(
        [Parameter(Mandatory=$false)]
        [switch]$Interactive = $true
    )
    
    $softwareList = @(
        @{Name = "Adobe Acrobat Reader"; Id = "Adobe.Acrobat.Reader.64-bit"; Selected = $true},
        @{Name = "Java 18 JRE"; Id = "EclipseAdoptium.Temurin.18.JRE"; Selected = $true},
        @{Name = "Mozilla Firefox"; Id = "Mozilla.Firefox"; Selected = $true},
        @{Name = "Google Chrome"; Id = "Google.Chrome"; Selected = $true},
        @{Name = "Microsoft Edge"; Id = "Microsoft.Edge"; Selected = $true},
        @{Name = "Microsoft OneDrive"; Id = "Microsoft.OneDrive"; Selected = $true},
        @{Name = "Microsoft Office"; Id = "Microsoft.Office"; Selected = $true},
        @{Name = "7-Zip"; Id = "7zip.7zip"; Selected = $true},
        @{Name = "Belgian eID Middleware"; Id = "BelgianGovernment.Belgium-eIDmiddleware"; Selected = $true},
        @{Name = "Belgian eID Viewer"; Id = "BelgianGovernment.eIDViewer"; Selected = $true}
    )
    
    # Check if winget is available
    $wingetExists = Test-CommandExists winget
    if (!$wingetExists) {
        Write-Log "Winget is not installed. Software installation will be skipped." -Level "WARNING"
        return $false
    }
    
    # Allow user to select software if in interactive mode
    if ($Interactive) {
        Write-Log "Software selection mode enabled" -Level "INFO"
        
        Write-Host ""
        Write-Host "=== Software Selection ===" -ForegroundColor Cyan
        Write-Host "Select which software to install. Enter the number(s) separated by commas, 'all' for everything, or 'none' to skip."
        Write-Host ""
        
        # Display software list with numbers
        for ($i = 0; $i -lt $softwareList.Count; $i++) {
            Write-Host "[$($i+1)] $($softwareList[$i].Name)" -ForegroundColor White
        }
        
        Write-Host ""
        $selection = Read-Host "Enter your selection"
        
        # Process user selection
        if ($selection -eq "none") {
            # Deselect all
            for ($i = 0; $i -lt $softwareList.Count; $i++) {
                $softwareList[$i].Selected = $false
            }
            Write-Log "User chose to skip all software installations" -Level "INFO"
        }
        elseif ($selection -ne "all") {
            # Deselect all first
            for ($i = 0; $i -lt $softwareList.Count; $i++) {
                $softwareList[$i].Selected = $false
            }
            
            # Select only user-chosen items
            $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() }
            foreach ($index in $selectedIndices) {
                if ($index -match '^\d+$' -and [int]$index -ge 1 -and [int]$index -le $softwareList.Count) {
                    $softwareList[[int]$index-1].Selected = $true
                    Write-Log "User selected: $($softwareList[[int]$index-1].Name)" -Level "INFO"
                }
            }
        }
        else {
            Write-Log "User selected all software for installation" -Level "INFO"
        }
    }
    
    # Filter only selected software
    $selectedSoftware = $softwareList | Where-Object { $_.Selected -eq $true }
    
    if ($selectedSoftware.Count -eq 0) {
        Write-Log "No software selected for installation" -Level "INFO"
        return $true
    }
    
    Write-Log "Starting software installation..." -Level "INFO"
    $installCount = 0
    $totalCount = $selectedSoftware.Count
    
    foreach ($software in $selectedSoftware) {
        Write-Log "Installing $($software.Name) ($($software.Id))..." -Level "INFO"
        
        Try {
            # Run winget with silent install
            $process = Start-Process -FilePath "winget" -ArgumentList "install -h -e --id $($software.Id)" -NoNewWindow -Wait -PassThru
            
            if ($process.ExitCode -eq 0) {
                Write-Log "Successfully installed $($software.Name)" -Level "SUCCESS"
                $installCount++
            }
            else {
                Write-Log "Failed to install $($software.Name) (Exit code: $($process.ExitCode))" -Level "WARNING"
            }
        }
        Catch {
            Write-Log "Error installing $($software.Name): $_" -Level "ERROR"
        }
    }
    
    Write-Log "Software installation complete. Installed $installCount of $totalCount packages." -Level "INFO"
    return ($installCount -eq $totalCount)
}

# Configure Start Menu
Function Set-StartMenu {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$StartBinPath
    )
    
    Write-Log "Configuring Start Menu..." -Level "INFO"
    $targetPath = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"
    
    Try {
        if (!(Test-Path -Path $targetPath)) {
            Write-Log "Start Menu target directory not found: $targetPath" -Level "WARNING"
            New-Item -Path $targetPath -ItemType Directory -Force | Out-Null
            Write-Log "Created Start Menu directory" -Level "INFO"
        }
        
        Copy-Item $StartBinPath -Destination $targetPath -Force
        Write-Log "Start Menu configuration applied successfully" -Level "SUCCESS"
        return $true
    }
    Catch {
        Write-Log "Failed to configure Start Menu: $_" -Level "ERROR"
        return $false
    }
}

# Function to refresh the desktop after registry changes
Function Update-Desktop {
    Param()
    
    Write-Log "Refreshing desktop to apply changes..." -Level "INFO"
    
    Try {
        # Restart Explorer to apply desktop icon changes
        $explorerProcess = Get-Process -Name explorer -ErrorAction SilentlyContinue
        
        if ($explorerProcess) {
            Write-Log "Restarting Windows Explorer..." -Level "INFO"
            Stop-Process -Name explorer -Force
            Start-Sleep -Seconds 1
            
            # Explorer should restart automatically, but we'll make sure
            if (!(Get-Process -Name explorer -ErrorAction SilentlyContinue)) {
                Start-Process explorer
            }
            
            Write-Log "Windows Explorer restarted successfully" -Level "SUCCESS"
            return $true
        }
        else {
            Write-Log "Windows Explorer process not found" -Level "WARNING"
            return $false
        }
    }
    Catch {
        Write-Log "Failed to refresh desktop: $_" -Level "ERROR"
        return $false
    }
}

# Configure Registry Settings
Function Set-RegistryConfigurations {
    Param (
        [Parameter(Mandatory=$false)]
        [switch]$Interactive = $true
    )
    
    Write-Log "Applying registry configurations..." -Level "INFO"
    
    # If no registry settings are defined, return success
    if ($Config.RegistrySettings.Count -eq 0) {
        Write-Log "No registry settings defined. Skipping." -Level "INFO"
        return $true
    }
    
    # Allow user to select registry settings if in interactive mode
    $selectedSettings = $Config.RegistrySettings
    
    if ($Interactive -and $Config.RegistrySettings.Count -gt 0) {
        Write-Host ""
        Write-Host "=== Registry Configuration ===" -ForegroundColor Cyan
        Write-Host "The following registry settings will be applied:" -ForegroundColor White
        Write-Host ""
        
        # Display registry settings with numbers
        for ($i = 0; $i -lt $Config.RegistrySettings.Count; $i++) {
            $setting = $Config.RegistrySettings[$i]
            Write-Host "[$($i+1)] $($setting.Path) -> $($setting.Name) = $($setting.Value) ($($setting.Type))" -ForegroundColor White
        }
        
        Write-Host ""
        $applySettings = Read-Host "Apply these registry settings? (Y/N)"
        
        if ($applySettings -ne "Y" -and $applySettings -ne "y") {
            Write-Log "User chose to skip registry configurations" -Level "INFO"
            return $true
        }
    }
    
    $successCount = 0
    $totalCount = $selectedSettings.Count
    $desktopIconsChanged = $false
    
    foreach ($setting in $selectedSettings) {
        Write-Log "Setting registry value: $($setting.Path) -> $($setting.Name) = $($setting.Value) ($($setting.Type))" -Level "INFO"
        
        Try {
            # Ensure the registry path exists
            if (!(Test-Path -Path $setting.Path)) {
                Write-Log "Creating registry path: $($setting.Path)" -Level "INFO"
                New-Item -Path $setting.Path -Force | Out-Null
            }
            
            # Set the registry value
            New-ItemProperty -Path $setting.Path -Name $setting.Name -Value $setting.Value -PropertyType $setting.Type -Force | Out-Null
            
            # Verify the setting was applied
            $verifyValue = Get-ItemProperty -Path $setting.Path -Name $setting.Name -ErrorAction SilentlyContinue
            
            if ($verifyValue -and $verifyValue.$($setting.Name) -eq $setting.Value) {
                Write-Log "Successfully applied registry setting: $($setting.Name)" -Level "SUCCESS"
                $successCount++
                
                # Check if this is a desktop icon setting
                if ($setting.Path -like "*HideDesktopIcons*") {
                    $desktopIconsChanged = $true
                }
            }
            else {
                Write-Log "Failed to verify registry setting: $($setting.Name)" -Level "WARNING"
            }
        }
        Catch {
            Write-Log "Error applying registry setting $($setting.Name): $_" -Level "ERROR"
        }
    }
    
    Write-Log "Registry configuration complete. Applied $successCount of $totalCount settings." -Level "INFO"
    
    # Refresh desktop if desktop icons were changed
    if ($desktopIconsChanged) {
        if ($Interactive) {
            $refreshDesktop = Read-Host "Desktop icon settings were changed. Refresh desktop now? (Y/N)"
            if ($refreshDesktop -eq "Y" -or $refreshDesktop -eq "y") {
                Update-Desktop
            }
            else {
                Write-Log "Desktop refresh skipped. Changes will apply after restart." -Level "INFO"
            }
        }
        else {
            # In non-interactive mode, refresh automatically
            Update-Desktop
        }
    }
    
    return ($successCount -eq $totalCount)
}

# Main function to orchestrate the installation process
Function Start-Installation {
    Param(
        [Parameter(Mandatory=$false)]
        [switch]$NonInteractive
    )
    
    Write-Log "=== VBC NV Workstation Setup Started ===" -Level "INFO"
    Write-Log "Script version: 1.1" -Level "INFO"
    Write-Log "Start time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level "INFO"
    
    # Step 1: Create installation directory
    $folderCreated = Initialize-InstallFolder
    if (!$folderCreated) {
        Write-Log "Failed to create installation folder. Aborting." -Level "ERROR"
        return
    }
    
    # Step 2: Install required software
    $softwareInstalled = Install-RequiredSoftware -Interactive:(!$NonInteractive)
    if (!$softwareInstalled) {
        Write-Log "Some software installations failed. Continuing with setup." -Level "WARNING"
    }
    
    # Step 3: Download and set wallpaper
    $wallpaperPath = "$($Config.InstallFolder)\bg-default.png"
    $wallpaperDownloaded = Get-FileWithProgress -Url $Config.WallpaperUrl -OutFile $wallpaperPath -ExpectedChecksum $Config.WallpaperChecksum
    
    if ($wallpaperDownloaded) {
        Set-Wallpaper -WallpaperPath $wallpaperPath
    }
    else {
        Write-Log "Wallpaper setup failed. Continuing with setup." -Level "WARNING"
    }
    
    # Step 4: Download and configure Start Menu
    $startBinPath = "$($Config.InstallFolder)\start2.bin"
    $startBinDownloaded = Get-FileWithProgress -Url $Config.StartBinUrl -OutFile $startBinPath -ExpectedChecksum $Config.StartBinChecksum
    
    if ($startBinDownloaded) {
        Set-StartMenu -StartBinPath $startBinPath
    }
    else {
        Write-Log "Start Menu configuration failed. Continuing with setup." -Level "WARNING"
    }
    
    # Step 5: Apply registry settings
    $registryConfigured = Set-RegistryConfigurations -Interactive:(!$NonInteractive)
    if (!$registryConfigured) {
        Write-Log "Some registry settings could not be applied. Continuing with setup." -Level "WARNING"
    }
    
    # Step 6: Finalize setup
    Write-Log "=== VBC NV Workstation Setup Completed ===" -Level "SUCCESS"
    Write-Log "End time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level "INFO"
    Write-Log "Log file location: $($Config.LogFile)" -Level "INFO"
    
    # Prompt for restart
    $restartChoice = Read-Host "Setup complete. Restart computer now? (Y/N)"
    if ($restartChoice -eq "Y" -or $restartChoice -eq "y") {
        Write-Log "Restarting computer..." -Level "INFO"
        Restart-Computer -Force
    }
    else {
        Write-Log "Restart skipped. Please restart your computer manually to complete setup." -Level "WARNING"
    }
}

# Main execution block
# This allows the script to be run via iex and still receive parameters
if ($MyInvocation.Line -match "iex|Invoke-Expression") {
    Write-Log "Script executed via Invoke-Expression" -Level "INFO"
    
    # Check if -NonInteractive was specified in the original command
    $nonInteractiveFlag = $false
    if ($MyInvocation.Line -match "-NonInteractive") {
        $nonInteractiveFlag = $true
        Write-Log "Non-interactive mode detected from command line" -Level "INFO"
    }
    
    # Run the installation with the appropriate mode
    Start-Installation -NonInteractive:$nonInteractiveFlag
}
else {
    # Normal execution with parameters passed directly to the script
    Start-Installation -NonInteractive:$NonInteractive
}
