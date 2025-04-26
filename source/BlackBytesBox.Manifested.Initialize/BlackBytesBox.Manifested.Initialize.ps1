function Register-LocalGalleryRepository {
    <#
    .SYNOPSIS
        Registers a local PowerShell repository for gallery modules.

    .DESCRIPTION
        This function ensures that the specified local repository folder exists, removes any existing
        repository with the given name, and registers the repository with a Trusted installation policy.

    .PARAMETER RepositoryPath
        The file system path to the local repository folder. Default is "$HOME/source/gallery".

    .PARAMETER RepositoryName
        The name to assign to the registered repository. Default is "LocalGallery".

    .EXAMPLE
        Register-LocalGalleryRepository
        Registers the local repository using the default path and repository name.

    .EXAMPLE
        Register-LocalGalleryRepository -RepositoryPath "C:\MyRepo" -RepositoryName "MyGallery"
        Registers the repository at "C:\MyRepo" with the name "MyGallery".
    #>
    [CmdletBinding()]
    [alias("rlgr")]
    param(
        [string]$RepositoryPath = "$HOME/source/gallery",
        [string]$RepositoryName = "LocalGallery"
    )

    # Normalize the repository path by replacing forward and backslashes with the platform's directory separator.
    $RepositoryPath = $RepositoryPath -replace '[/\\]', [System.IO.Path]::DirectorySeparatorChar

    # Ensure the local repository folder exists; if not, create it.
    if (-not (Test-Path -Path $RepositoryPath)) {
        New-Item -ItemType Directory -Path $RepositoryPath | Out-Null
    }

    # If a repository with the specified name exists, unregister it.
    if (Get-PSRepository -Name $RepositoryName -ErrorAction SilentlyContinue) {
        Write-Host "Repository '$RepositoryName' already exists. Removing it." -ForegroundColor Yellow
        Unregister-PSRepository -Name $RepositoryName
    }

    # Register the local PowerShell repository with a Trusted installation policy.
    Register-PSRepository -Name $RepositoryName -SourceLocation $RepositoryPath -InstallationPolicy Trusted

    Write-Host "Local repository '$RepositoryName' registered at: $RepositoryPath" -ForegroundColor Green
}

function Update-ManifestModuleVersion {
    <#
    .SYNOPSIS
        Updates the ModuleVersion in a PowerShell module manifest (psd1) file.

    .DESCRIPTION
        This function reads a PowerShell module manifest file as text, uses a regular expression to update the
        ModuleVersion value while preserving the file's comments and formatting, and writes the updated content back
        to the file. If a directory path is supplied, the function recursively searches for the first *.psd1 file and uses it.

    .PARAMETER ManifestPath
        The file or directory path to the module manifest (psd1) file. If a directory is provided, the function will
        search recursively for the first *.psd1 file.

    .PARAMETER NewVersion
        The new version string to set for the ModuleVersion property.

    .EXAMPLE
        PS C:\> Update-ManifestModuleVersion -ManifestPath "C:\projects\MyDscModule" -NewVersion "2.0.0"
        Updates the ModuleVersion of the first PSD1 manifest found in the given directory to "2.0.0".
    #>
    [CmdletBinding()]
    [alias("ummv")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$NewVersion
    )

    # Check if the provided path exists
    if (-not (Test-Path $ManifestPath)) {
        throw "The path '$ManifestPath' does not exist."
    }

    # If the path is a directory, search recursively for the first *.psd1 file.
    $item = Get-Item $ManifestPath
    if ($item.PSIsContainer) {
        $psd1File = Get-ChildItem -Path $ManifestPath -Filter *.psd1 -Recurse | Select-Object -First 1
        if (-not $psd1File) {
            throw "No PSD1 manifest file found in directory '$ManifestPath'."
        }
        $ManifestPath = $psd1File.FullName
    }

    Write-Verbose "Using manifest file: $ManifestPath"

    # Read the manifest file content as text using .NET method.
    $content = [System.IO.File]::ReadAllText($ManifestPath)

    # Define the regex pattern to locate the ModuleVersion value.
    $pattern = "(?<=ModuleVersion\s*=\s*')[^']+(?=')"

    # Replace the current version with the new version using .NET regex.
    $updatedContent = [System.Text.RegularExpressions.Regex]::Replace($content, $pattern, $NewVersion)

    # Write the updated content back to the manifest file.
    [System.IO.File]::WriteAllText($ManifestPath, $updatedContent)
}

function Update-ModuleIfNewer {
    <#
    .SYNOPSIS
        Installs or updates a module from a repository only if a newer version is available.

    .DESCRIPTION
        This function uses Find-Module to search for a module (default repository is PSGallery) and compares the
        remote version with the locally installed version (if any) using Get-InstalledModule. If the module is not installed
        or the remote version is newer, it then installs the module using Install-Module. This prevents forcing a download
        when the installed module is already up to date.

    .PARAMETER ModuleName
        The name of the module to check and install/update.

    .PARAMETER Repository
        The repository from which to search for the module. Defaults to 'PSGallery'.

    .EXAMPLE
        PS C:\> Update-ModuleIfNewer -ModuleName 'STROM.NANO.PSWH.CICD'
        Searches PSGallery for the module 'STROM.NANO.PSWH.CICD' and installs it only if it is not installed or if a newer version is available.
    #>
    [CmdletBinding()]
    [alias("umn")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,

        [Parameter(Mandatory = $false)]
        [string]$Repository = 'PSGallery'
    )

    try {
        Write-Verbose "Searching for module '$ModuleName' in repository '$Repository'..."
        $remoteModule = Find-Module -Name $ModuleName -Repository $Repository -ErrorAction Stop

        if (-not $remoteModule) {
            Write-Error "Module '$ModuleName' not found in repository '$Repository'."
            return
        }

        $remoteVersion = [version]$remoteModule.Version

        # Check if the module is installed locally.
        $localModule = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue

        if ($localModule) {
            $localVersion = [version]$localModule.Version
            if ($remoteVersion -gt $localVersion) {
                Write-Host "A newer version ($remoteVersion) is available (local version: $localVersion). Installing update..."
                Install-Module -Name $ModuleName -Repository $Repository -Force
            }
            else {
                Write-Host "The installed module ($localVersion) is up to date."
            }
        }
        else {
            Write-Host "Module '$ModuleName' is not installed. Installing version $remoteVersion..."
            Install-Module -Name $ModuleName -Repository $Repository -Force
        }
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

function Remove-OldModuleVersions {
    <#
    .SYNOPSIS
        Removes older versions of an installed PowerShell module, keeping only the latest version.

    .DESCRIPTION
        This function retrieves all installed versions of a specified module, sorts them by version in descending
        order (so that the latest version is first), and removes all versions except the latest one.
        It helps clean up local installations accumulated from repeated updates.

    .PARAMETER ModuleName
        The name of the module for which to remove older versions. Only versions beyond the latest one are removed.

    .EXAMPLE
        PS C:\> Remove-OldModuleVersions -ModuleName 'STROM.NANO.PSWH.CICD'
        Removes all installed versions of 'STROM.NANO.PSWH.CICD' except for the latest version.
    #>
    [CmdletBinding()]
    [alias("romv")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        # Retrieve all installed versions of the module.
        $installedModules = Get-InstalledModule -Name $Name -AllVersions -ErrorAction SilentlyContinue

        if (-not $installedModules) {
            Write-Host "No installed module found with the name '$Name'." -ForegroundColor Yellow
            return
        }

        # Sort installed versions descending; latest version comes first.
        $sortedModules = $installedModules | Sort-Object -Property Version -Descending

        # Retain the latest version (first item) and select all older versions.
        $latestModule = $sortedModules[0]
        $oldModules = $sortedModules | Select-Object -Skip 1

        if (-not $oldModules) {
            Write-Host "No older versions of '$Name' to remove." -ForegroundColor Green
            return
        }

        foreach ($module in $oldModules) {
            Write-Host "Removing $Name version $($module.Version)..." -ForegroundColor Cyan
            Uninstall-Module -Name $Name -RequiredVersion $module.Version -Force
        }

        Write-Host "Cleaned up '$Name'; latest ($($latestModule.Version)) kept."
        
    }
    catch {
        Write-Error "An error occurred while removing old versions: $_"
    }
}

function Install-UserModule {
    <#
    .SYNOPSIS
        Installs one or more modules for the *current* user.
    
    .DESCRIPTION
        Thin wrapper around Install‑Module that forces `-Scope CurrentUser`
        but otherwise behaves just like the original cmdlet.
    
    .PARAMETER Name
        The module name(s) to install.
    
    .PARAMETER Force
        Suppresses all prompts, mirroring Install‑Module’s -Force switch.
    
    .EXAMPLE
        Install-UserModule -Name Pester -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string[]] $Name,
    
        [switch] $Force
    )
    
    # Inject / override the scope and forward everything else
    $PSBoundParameters['Scope'] = 'CurrentUser'
    Install-Module @PSBoundParameters
}   
    
function Initialize-DotNet {
    <#
    .SYNOPSIS
        Installs specified .NET channels and sets environment variables for both the current session and the user profile.

    .DESCRIPTION
        This function performs the following actions:
        
          1. For each provided channel (defaulting to 8.0 and 9.0 if none are specified):
             - Sets TLS12 as the security protocol.
             - Downloads and executes the dotnet-install.ps1 script using Invoke-WebRequest with RawContent,
               and passes the -channel parameter.
          
          2. Sets the DOTNET_ROOT environment variable to "$HOME\.dotnet" for the user and current session.
          3. Updates the user's PATH environment variable to include both DOTNET_ROOT and the tools folder
             ("$HOME\.dotnet\tools") and updates the current session PATH accordingly.

    .PARAMETER Channels
        An array of .NET channels to install. If omitted, the function defaults to installing channels 8.0 and 9.0.

    .EXAMPLE
        PS C:\> Initialize-DotNet
        Installs .NET channels 8.0 and 9.0, and configures the environment variables for immediate and persistent use.

    .EXAMPLE
        PS C:\> Initialize-DotNet -Channels @("2.1","2.2","3.0","3.1","5.0", "6.0", "7.0", "8.0", "9.0")
        Installs the specified .NET channels and configures the environment variables.
    #>
    [CmdletBinding()]
    [alias("idot")]
    param(
        [string[]]$Channels = @("8.0", "9.0")
    )

    $dotnetInstallUrl = 'https://dot.net/v1/dotnet-install.ps1'

    foreach ($channel in $Channels) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Host "Installing .NET channel $channel..." -ForegroundColor Cyan
        & ([scriptblock]::Create((Invoke-WebRequest -UseBasicParsing $dotnetInstallUrl))) -channel $channel -InstallDir "$HOME\.dotnet"
    }

    # Set DOTNET_ROOT environment variable for both persistent and current session.
    $dotnetRoot = "$HOME\.dotnet"
    [Environment]::SetEnvironmentVariable('DOTNET_ROOT', $dotnetRoot, 'User')
    $env:DOTNET_ROOT = $dotnetRoot
    Write-Host "DOTNET_ROOT set to $dotnetRoot" -ForegroundColor Green
    [Environment]::SetEnvironmentVariable('DOTNET_CLI_TELEMETRY_OPTOUT', 'true', 'User')
    $env:DOTNET_CLI_TELEMETRY_OPTOUT = 'true'
    Write-Host "DOTNET_CLI_TELEMETRY_OPTOUT set to true" -ForegroundColor Green

    # Define the tools folder.
    $toolsFolder = "$dotnetRoot\tools"

    # Update PATH to include DOTNET_ROOT and the tools folder for persistent storage.
    $currentPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    $pathsToAdd = @()

    if (-not $currentPath.ToLower().Contains($dotnetRoot.ToLower())) {
        $pathsToAdd += $dotnetRoot
    }
    if (-not $currentPath.ToLower().Contains($toolsFolder.ToLower())) {
        $pathsToAdd += $toolsFolder
    }
    if ($pathsToAdd.Count -gt 0) {
        $newPath = "$currentPath;" + ($pathsToAdd -join ';')
        [Environment]::SetEnvironmentVariable('PATH', $newPath, 'User')
        # Also update the current session's PATH immediately.
        $env:PATH = $newPath
        Write-Host "PATH updated to include: $($pathsToAdd -join ', ')" -ForegroundColor Green
    }
    else {
        Write-Host "PATH already contains DOTNET_ROOT and tools folder." -ForegroundColor Yellow
    }
}

function Initialize-NugetRepositoryDotNet {
    <#
    .SYNOPSIS
        Initializes a NuGet package source using the dotnet CLI.

    .DESCRIPTION
        This function manages a single NuGet source using the dotnet CLI. It retrieves the currently registered
        sources via 'dotnet nuget list source' and checks if the provided source (by Name and Location) exists.
        If the source is not found, it registers it. If the source is found but is marked as [Disabled],
        it removes and then re-adds the source as enabled.
        Additionally, if the Location is a local path (not a URL), it ensures the directory exists by creating it if necessary.

    .EXAMPLE
        Initialize-NugetRepositoryDotNet -Name "nuget.org" -Location "https://api.nuget.org/v3/index.json"
        This will verify that the NuGet source for nuget.org is registered and enabled.
    #>
    [CmdletBinding()]
    [alias("inugetx")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Location
    )

    # Check if the Location is a URL; if not, treat it as a local directory.
    if ($Location -notmatch '^https?://') {
        $Location = $Location -replace '[/\\]', [System.IO.Path]::DirectorySeparatorChar
        Write-Host "Provided Location '$Location' is a local path." -ForegroundColor Cyan
        if (-not (Test-Path $Location)) {
            Write-Host "Local path '$Location' does not exist. Creating directory." -ForegroundColor Cyan
            New-Item -ItemType Directory -Path $Location | Out-Null
        }
    }

    Write-Host "Retrieving registered NuGet sources using dotnet CLI..." -ForegroundColor Cyan
    $listOutput = dotnet nuget list source 2>&1
    $lines = $listOutput -split "`n"

    $foundIndex = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match [regex]::Escape($Location)) {
            $foundIndex = $i
            break
        }
    }

    if ($foundIndex -ne $null) {
        # Assume the preceding line contains the name and status, e.g., " 1.  nuget.org [Enabled]"
        $statusLine = if ($foundIndex -gt 0) { $lines[$foundIndex - 1] } else { "" }
        if ($statusLine -match '^\s*\d+\.\s*(?<Name>\S+)\s*\[(?<Status>\w+)\]') {
            $registeredName = $Matches["Name"]
            $status = $Matches["Status"]
            if ($status -eq "Disabled") {
                Write-Host "Source '$registeredName' ($Location) is disabled. Removing and re-adding it as enabled." -ForegroundColor Yellow
                dotnet nuget remove source $registeredName
                Write-Host "Adding source '$Name' with URL '$Location'." -ForegroundColor Green
                dotnet nuget add source $Location --name $Name
            }
            else {
                Write-Host "Source '$registeredName' with URL '$Location' is already registered and enabled. Skipping." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Could not parse status for source with URL '$Location'. Skipping." -ForegroundColor Red
        }
    }
    else {
        Write-Host "Source '$Name' not found. Registering it." -ForegroundColor Green
        dotnet nuget add source $Location --name $Name
    }
}

function Initialize-NugetRepositories {
    <#
    .SYNOPSIS
        Initializes the default NuGet package sources.

    .DESCRIPTION
        This function registers the default NuGet package sources if they are not already present.
        It uses enhanced logic: if a repository with a matching URL exists but is not trusted,
        it will be re-registered with the Trusted flag. If the repository exists and is already trusted,
        it is skipped.

    .EXAMPLE
        Init-NugetRepositorys
        Initializes and registers the default NuGet package sources, ensuring they are trusted.
    #>
    [CmdletBinding()]
    [alias("inuget")]
    param()
    # Define the default NuGet repository sources.
    $defaultSources = @(
        [PSCustomObject]@{ Name = "nuget.org";         Location = "https://api.nuget.org/v3/index.json" },
        [PSCustomObject]@{ Name = "int.nugettest.org"; Location = "https://apiint.nugettest.org/v3/index.json" }
    )

    # Retrieve the currently registered NuGet package sources.
    $existingSources = Get-PackageSource -ProviderName NuGet -ErrorAction SilentlyContinue

    foreach ($source in $defaultSources) {
        $found = $existingSources | Where-Object { $_.Location -eq $source.Location }
        if ($found) {
            # Check if the found source is trusted.
            if (-not $($found.IsTrusted)) {
                Write-Host "Repository '$($source.Name)' exists but is not trusted. Updating trust setting." -ForegroundColor Yellow
                # Unregister the untrusted source and re-register it with the Trusted flag.
                Unregister-PackageSource -Name $found.Name -ProviderName NuGet -Force -ErrorAction SilentlyContinue
                Register-PackageSource -Name $source.Name -Location $source.Location -ProviderName NuGet -Trusted
            }
            else {
                Write-Host "Repository '$($source.Name)' with URL '$($source.Location)' is already registered and trusted. Skipping." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Registering repository '$($source.Name)' with URL '$($source.Location)'." -ForegroundColor Green
            Register-PackageSource -Name $source.Name -Location $source.Location -ProviderName NuGet -Trusted | Out-Null
        }
    }
}

function Test-IsWindows {
    <#
    .SYNOPSIS
        Returns $True if PowerShell is running on Windows (any edition or version).

    .DESCRIPTION
        Combines .NET RuntimeInformation (if present) with a PlatformID fallback
        to detect Windows accurately across PowerShell Desktop and Core.

    .OUTPUTS
        Boolean: $True on Windows, $False otherwise.
    #>
    [alias("iswin")]
    param()
    # Determine if the RuntimeInformation type exists
    if ([Type]::GetType('System.Runtime.InteropServices.RuntimeInformation', $false)) {
        # Use cross-platform API
        return [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
            [System.Runtime.InteropServices.OSPlatform]::Windows
        )
    }
    else {
        # Fallback for older Windows PowerShell: check PlatformID
        return (
            [Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
        )
    }
}

<#
.SYNOPSIS
    Writes a timestamped, color‑coded inline log entry to the console, optionally appends to a daily log file, and optionally returns log details as JSON.

.DESCRIPTION
    Formats messages with a high-precision timestamp, log-level abbreviation, and caller identifier. Color-codes console output by severity, can overwrite the previous line, and can append to a per-process daily log file.
    Use -ReturnJson to emit a JSON representation of the log details instead of returning nothing.

.PARAMETER Level
    The log level. Valid values: Verbose, Debug, Information, Warning, Error, Critical.

.PARAMETER MinLevel
    Minimum level to write to the console. Messages below this level are suppressed. Default: Information.

.PARAMETER FileMinLevel
    Minimum level to append to the log file. Messages below this level are skipped. Default: Verbose.

.PARAMETER Template
    The message template, using placeholders like {Name}.

.PARAMETER Params
    Values for each placeholder in Template. Either a hashtable or an ordered object array.

.PARAMETER UseBackColor
    Switch to enable background coloring in the console.

.PARAMETER Overwrite
     Switch to overwrite the previous console entry rather than writing a new line.

.PARAMETER InitialWrite
    Switch to output an initial blank line instead of attempting to overwrite on the first call when using -Overwrite.

.PARAMETER FileAppName
    When set, enables file logging under:
      %LOCALAPPDATA%\Write-LogInline\<FileAppName>\<yyyy-MM-dd>_<PID>.log

.PARAMETER ReturnJson
    Switch to return the log details as a JSON-formatted string; otherwise, no output.

.EXAMPLE
    # Write a green "Hello, World!" message to the console
    Write-LogInline -Level Information `
                   -Template "{greeting}, {user}!" `
                   -Params @{ greeting = "Hello"; user = "World" }

.EXAMPLE
    # Using defaults plus -ReturnJson
    $WriteLogInlineDefaults = @{
        FileMinLevel  = 'Verbose'
        MinLevel      = 'Information'
        UseBackColor  = $false
        Overwrite     = $true
        FileAppName   = 'testing'
        ReturnJson    = $false
    }

    Write-LogInline -Level Verbose `
                   -Template "{hello}-{world} number {num} at {time}!" `
                   -Params "Hello","World",1,1.2 `
                   @WriteLogInlineDefaults

.NOTES
    Requires PowerShell 5.0 or later.
#>
function Write-LogInline {
    [CmdletBinding()]
    [alias("wlog")]
    param(
        [ValidateSet('Verbose','Debug','Information','Warning','Error','Critical')][string]$Level,
        [ValidateSet('Verbose','Debug','Information','Warning','Error','Critical')][string]$MinLevel       = 'Information',
        [ValidateSet('Verbose','Debug','Information','Warning','Error','Critical')][string]$FileMinLevel  = 'Verbose',
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()] [string]$Template,
        [object]$Params,
        [switch]$UseBackColor,
        [switch]$Overwrite,
        [switch]$InitialWrite,
        [string]$FileAppName,
        [switch]$ReturnJson
    )

    # Normalize any non-hashtable, non-array to a one‐item array
    if ($Params -isnot [hashtable] -and $Params -isnot [object[]]) {
        $Params = @($Params)
    }

    # Now enforce flatness on arrays
    if ($Params -is [object[]] -and ($Params |
    Where-Object { $_ -is [System.Collections.IEnumerable] -and -not ($_ -is [string]) }
    )) {
        throw "Parameter -Params array must be flat (no nested collections)."
    }

    # ANSI escape
    $esc = [char]27
    if (-not $script:WLI_Caller) {
        $script:WLI_Caller = if ($MyInvocation.PSCommandPath) { Split-Path -Leaf $MyInvocation.PSCommandPath } else { 'Console' }
    }
    $caller = $script:WLI_Caller

    # Level maps
    $levelValues = @{ Verbose=0; Debug=1; Information=2; Warning=3; Error=4; Critical=5 }
    $abbrMap      = @{ Verbose='VRB'; Debug='DBG'; Information='INF'; Warning='WRN'; Error='ERR'; Critical='CRT' }

    $writeConsole = $levelValues[$Level] -ge $levelValues[$MinLevel]
    $writeToFile  = $FileAppName -and ($levelValues[$Level] -ge $levelValues[$FileMinLevel])
    if (-not ($writeConsole -or $writeToFile)) { return }

    # File path init
    if ($writeToFile) {
        $os = [int][System.Environment]::OSVersion.Platform
        switch ($os) {
            2 { $base = $env:LOCALAPPDATA } # Win32NT
            4 { $base = Join-Path $env:HOME ".local/share" } # Unix
            6 { $base = Join-Path $env:HOME ".local/share" } # MacOSX
            default { throw "Unsupported OS platform: $os" }
        }
        $root = Join-Path $base "Write-LogInline/$FileAppName"

        if (-not (Test-Path $root)) { New-Item -Path $root -ItemType Directory | Out-Null }
        $date    = Get-Date -Format 'yyyy-MM-dd'
        $logPath = Join-Path $root "${date}_${PID}.log"
    }

    # Timestamp and render
    $timeEntry = Get-Date
    $timeStr   = $timeEntry.ToString('yyyy-MM-dd HH:mm:ss:ff')
    $plMatches = [regex]::Matches($Template, '{(?<name>\w+)}')
    $keys      = $plMatches | ForEach-Object { $_.Groups['name'].Value } | Select-Object -Unique
    $wasHash    = $Params -is [hashtable]
    $paramArray = @($Params)

    if (-not $wasHash -and $paramArray.Count -lt $keys.Count) {
        throw "Insufficient parameters: expected $($keys.Count), received $($paramArray.Count)"
    }

    $keys = @($keys)
    if ($wasHash) {
        $map = $Params
    } else {
        $map = @{}
        for ($i = 0; $i -lt $keys.Count; $i++) { $map[$keys[$i]] = $paramArray[$i] }  # CHANGED: use paramArray
    }

    # Fix: cast null to empty string, avoid boolean -or misuse
    $msg = $Template
    foreach ($k in $keys) {
        $escName = [regex]::Escape($k)
        $msg = $msg -replace "\{$escName\}", [string]$map[$k]
    }
    $rawLine = "[$timeStr $($abbrMap[$Level])][$caller] $msg"

    # Write to file
    if ($writeToFile) { $rawLine | Out-File -FilePath $logPath -Append -Encoding UTF8 }

    # Console output
    if ($writeConsole) {
        if ($InitialWrite) {
            # Initial invocation: write a blank line instead of overwriting
            Write-Host ""
        }
        if ($Overwrite) {
            for ($i = 0; $i -lt $script:WLI_LastLines; $i++) {
                Write-Host -NoNewline ($esc + '[1A' + "`r" + $esc + '[K')
            }
        }
        Write-Host -NoNewline ($esc + '[?25l')

        # Color maps
        $levelMap = @{
            Verbose     = @{ Abbrev='VRB'; Fore='DarkGray' }
            Debug       = @{ Abbrev='DBG'; Fore='Cyan'     }
            Information = @{ Abbrev='INF'; Fore='Green'    }
            Warning     = @{ Abbrev='WRN'; Fore='Yellow'   }
            Error       = @{ Abbrev='ERR'; Fore='Red'      }
            Critical    = @{ Abbrev='CRT'; Fore='White'; Back='DarkRed' }
        }
        $typeColorMap = @{
            'System.String'   = 'Green';   'System.DateTime' = 'Yellow'
            'System.Int32'    = 'Cyan';    'System.Int64'     = 'Cyan'
            'System.Double'   = 'Blue';    'System.Decimal'   = 'Blue'
            'System.Boolean'  = 'Magenta'; 'Default'          = 'White'
            'System.Version'  = 'Magenta'; 'Microsoft.PackageManagement.Internal.Utility.Versions.FourPartVersion' = 'Magenta'
            'Microsoft.PowerShell.ExecutionPolicy' = 'Magenta'
            'System.Management.Automation.ActionPreference' = 'Green'
        }
        $staticFore = 'White'; $staticBack = 'Black'
        function Write-Colored { param($Text,$Fore,$Back) if ($UseBackColor -and $Back) { Write-Host -NoNewline $Text -ForegroundColor $Fore -BackgroundColor $Back } else { Write-Host -NoNewline $Text -ForegroundColor $Fore } }

        # Header
        $entry = $levelMap[$Level]
        $tag   = $entry.Abbrev
        if ($entry.ContainsKey('Back')) {
            $lvlBack = $entry.Back
        } elseif ($UseBackColor) {
            $lvlBack = $staticBack
        } else {
            $lvlBack = $null
        }
        Write-Colored '[' $staticFore $staticBack; Write-Colored $timeStr $staticFore $staticBack; Write-Colored ' ' $staticFore $staticBack
        if ($lvlBack) { Write-Host -NoNewline $tag -ForegroundColor $entry.Fore -BackgroundColor $lvlBack } else { Write-Host -NoNewline $tag -ForegroundColor $entry.Fore }
        Write-Colored '] [' $staticFore $staticBack; Write-Colored $caller $staticFore $staticBack; Write-Colored '] ' $staticFore $staticBack

        # Message parts
        $pos = 0
        foreach ($m in $plMatches) {
            if ($m.Index -gt $pos) {
                Write-Colored $Template.Substring($pos, $m.Index - $pos) $staticFore $staticBack
            }
            $val = $map[$m.Groups['name'].Value]
            $t   = $val.GetType().FullName

            if ($typeColorMap.ContainsKey($t)) {
                $f = $typeColorMap[$t]
            } else {
                $f = $typeColorMap['Default']
            }

            if ($UseBackColor) {
                $b = $staticBack
            } else {
                $b = $null
            }

            Write-Colored $val $f $b
            $pos = $m.Index + $m.Length
        }

        if ($pos -lt $Template.Length) {
            if ($UseBackColor) {
                $b = $staticBack
            } else {
                $b = $null
            }
            Write-Colored $Template.Substring($pos) $staticFore $b
        }

        Write-Host ''
        Write-Host -NoNewline ($esc + '[?25h')

        try {
            $width = $Host.UI.RawUI.WindowSize.Width
        } catch {
            $width = 80
        }

        $script:WLI_LastLines = [math]::Ceiling($rawLine.Length / ($width - 1))
    }

    # Return JSON
    $output = [PSCustomObject]@{
        DateTime   = $timeEntry
        PID        = $PID
        Level      = $Level
        Template   = $Template
        Message    = $msg
        Parameters = $map
    }

    # Return JSON only if requested
    if ($ReturnJson) {
        return $output | ConvertTo-Json -Depth 5
    }
}


<#
.SYNOPSIS
    Temporarily enables script and module execution by setting a permissive execution policy for the CurrentUser, while capturing the original policy.
.DESCRIPTION
    Retrieves the CurrentUser execution policy. If it's not already permissive, sets it to RemoteSigned. Returns the original policy for later restoration.
.EXAMPLE
    # Capture and enable script execution temporarily
    $originalPolicy = Enable-TemporaryUserScriptExecution
#>
function Enable-TemporaryUserScriptExecution {
    [CmdletBinding()]
    [alias("etse")]
    param()
    # Constant list of policies that allow scripts/modules
    $allowedPolicies = @('RemoteSigned','Unrestricted','Bypass')

    $WriteLogInlineDefaults = @{
        FileMinLevel  = 'Error'
        MinLevel      = 'Information'
        UseBackColor  = $false
        Overwrite     = $false
        FileAppName   = $null
        ReturnJson    = $false
    }

    # Capture the current CurrentUser policy
    $originalPolicy = Get-ExecutionPolicy -Scope CurrentUser

            # Log the invocation attempt
    
    try {
        Write-LogInline -Level Information -Template "Attempting to enable temporary script execution for CurrentUser scope..." @WriteLogInlineDefaults

        Write-LogInline -Level Information -Template "Current CurrentUser policy is '{0}'." -Params $originalPolicy @WriteLogInlineDefaults

        if ($allowedPolicies -contains $originalPolicy) {
            Write-LogInline -Level Information -Template "Policy is already permissive; no change needed." @WriteLogInlineDefaults
        }
        else {
            $targetPolicy = $allowedPolicies[0]  # RemoteSigned
            Write-LogInline -Level Information -Template "Temporarily setting CurrentUser execution policy to '{0}'." -Params $targetPolicy @WriteLogInlineDefaults
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy $targetPolicy -Force
            Write-LogInline -Level Information -Template "Execution policy now '{0}' for temporary script execution." -Params $targetPolicy  @WriteLogInlineDefaults
        }
    }
    catch {
        Write-LogInline -Level Error -Template "Error enabling temporary script execution: $_" @WriteLogInlineDefaults
        throw
    }

    # Return the original for restoration
    return $originalPolicy
}


<#
.SYNOPSIS
    Restores a previously captured CurrentUser execution policy to its original state, if needed.
.DESCRIPTION
    Compares the CurrentUser execution policy with the given original. If they differ, restores the policy. Otherwise logs that no change is needed.
.EXAMPLE
    # Restore policy after temporary change
    Restore-OriginalUserScriptExecution -Policy $originalPolicy
#>
function Restore-OriginalUserScriptExecution {
    [CmdletBinding()]
    [alias("rotse")]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Policy
    )

    $WriteLogInlineDefaults = @{
        FileMinLevel  = 'Error'
        MinLevel      = 'Information'
        UseBackColor  = $false
        Overwrite     = $false
        FileAppName   = $null
        ReturnJson    = $false
    }
    
    # Get the current policy to decide if restoration is needed
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser

    try {
        if ($currentPolicy -eq $Policy) { 
            Write-LogInline -Level Information -Template "CurrentUser policy restore is already '{0}' (desired was '{1}'); no restore needed." ` -Params $currentPolicy, $Policy @WriteLogInlineDefaults

        }
        else {
            Write-LogInline -Level Information -Template "Restoring CurrentUser execution policy to '{0}'." -Params $currentPolicy @WriteLogInlineDefaults
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy $currentPolicy -Force
            $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
            Write-LogInline -Level Information -Template "Execution policy restored to '{0}'." -Params $currentPolicy @WriteLogInlineDefaults
        }
    }
    catch {
        Write-LogInline -Level Error -Template "Failed to restore execution policy: $_" @WriteLogInlineDefaults
        throw
    }
}

function Invoke-IsolatedScript {
    <#
    .SYNOPSIS
        Executes a given PowerShell script block in an isolated shell and returns output and errors.

    .DESCRIPTION
        This function executes the provided PowerShell script block in a completely isolated and fresh PowerShell environment.
        It captures standard output and errors internally and returns them as structured results.

        Useful when script execution must avoid interference from existing loaded modules or persistent states.

    .PARAMETER ScriptBlock
        The PowerShell script block containing commands to execute.

    .EXAMPLE
        $result = Invoke-IsolatedScript -ScriptBlock { Remove-OldModuleVersions -Name 'Example.Module' }
        if ($result.Output) { $result.Output | ForEach-Object { Write-Host $_ } }
        if ($result.Errors) { $result.Errors | ForEach-Object { Write-Error $_ } }

    .RETURNS
        [PSCustomObject] containing Output and Errors arrays.
    #>
    [CmdletBinding()]
    [alias("iis")]
    param (
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$ScriptBlock
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command $($ScriptBlock.ToString())"
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $process = [System.Diagnostics.Process]::Start($psi)
    $output = $process.StandardOutput.ReadToEnd()
    $errorOutput = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    # Process and structure output
    $outputLines = $output -split "`n" | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne '' }
    $errorLines = $errorOutput -split "`n" | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne '' }

    return [PSCustomObject]@{
        Output = $outputLines
        Errors = $errorLines
    }
}




