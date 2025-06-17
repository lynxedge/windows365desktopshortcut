# Creates a Desktop Shortcut for the Windows App (aka Windows 365 app)
# It will find the latest verison of the Windows App and create the Desktop shortcut for that

# Define the Shortcut Location
$shortcutInstallPath = "C:\Users\vithinclient\Desktop"

# Define the shortcut name
$ShortcutName = "Windows App.lnk"

# Define common install paths
$InstallPaths = @(
    "C:\Program Files\WindowsApps"
)

# Function to get the latest version path
function Get-LatestVersionPath {
    param (
        [string]$basePath,
        [string]$appName
    )
    $latestVersionPath = $null
    $latestVersion = [version]::new(0,0,0,0)
    if (Test-Path $basePath) {
        Get-ChildItem -Path $basePath -Directory | ForEach-Object {
            if ($_.Name -match $appName) {
                $versionString = $_.Name -replace '^.*_([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)_.*$', '$1'
                $version = [version]::new($versionString)
                if ($version -gt $latestVersion) {
                    $latestVersion = $version
                    $latestVersionPath = $_.FullName
                }
            }
        }
    }
    return $latestVersionPath
}

# Define the app name pattern
$appNamePattern = "Windows365"

# Search for the latest version path
$latestPath = $null
foreach ($path in $InstallPaths) {
    $latestPath = Get-LatestVersionPath -basePath $path -appName $appNamePattern
    if ($latestPath) { break }
}

if ($latestPath) {
    
    #Write-Output "Found the app in:" + $latestPath


    # Define the target executable
    $TargetPath = Join-Path -Path $latestPath -ChildPath "Windows365.exe"

    Write-Output "Creating Shortcut for:" + $TargetPath

    # Define the shortcut path
    $ShortcutPath = Join-Path -Path $shortcutInstallPath -ChildPath "$ShortcutName"
	#$ShortcutPath = $shortcutInstallPath

    # Create WScript.Shell COM object
    $WScriptShell = New-Object -ComObject WScript.Shell

    # Create the shortcut
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.WorkingDirectory = Split-Path $TargetPath
    $Shortcut.WindowStyle = 1
    $Shortcut.Description = "Launch Windows App"
    $Shortcut.Save()

    Write-Output "Shortcut created on Desktop for Windows App."
} else {
    Write-Output "Windows 365 App not found in the expected directories."
}
