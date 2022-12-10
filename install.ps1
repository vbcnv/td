#Requires -RunAsAdministrator

Function Test-CommandExists
{
    Param ($command)
 
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'stop'
 
    Try { if (Get-Command $command) { $true } }
    Catch { $false }
    Finally { $ErrorActionPreference = $oldPreference }
}

$exists = Test-CommandExists winget

if ($exists) {
    winget install -e --id Adobe.Acrobat.Reader.64-bit
    winget install -e --id EclipseAdoptium.Temurin.18.JRE
    winget install -e --id Mozilla.Firefox
    winget install -e --id Google.Chrome
    winget install -e --id Microsoft.Edge
    winget install -e --id Microsoft.OneDrive
    winget install -e --id Microsoft.Office
    winget install -e --id ZeroTier.ZeroTierOne
    winget install -e --id 7zip.7zip
    winget install -e --id BelgianGovernment.Belgium-eIDmiddleware
    winget install -e --id BelgianGovernment.eIDViewer
}
