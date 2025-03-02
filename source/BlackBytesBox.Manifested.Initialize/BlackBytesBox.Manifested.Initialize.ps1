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
        [string]$ModuleName
    )

    try {
        # Retrieve all installed versions of the module.
        $installedModules = Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue

        if (-not $installedModules) {
            Write-Host "No installed module found with the name '$ModuleName'." -ForegroundColor Yellow
            return
        }

        # Sort installed versions descending; latest version comes first.
        $sortedModules = $installedModules | Sort-Object -Property Version -Descending

        # Retain the latest version (first item) and select all older versions.
        $latestModule = $sortedModules[0]
        $oldModules = $sortedModules | Select-Object -Skip 1

        if (-not $oldModules) {
            Write-Host "Only one version of '$ModuleName' is installed. Nothing to remove." -ForegroundColor Green
            return
        }

        foreach ($module in $oldModules) {
            Write-Host "Removing $ModuleName version $($module.Version)..." -ForegroundColor Cyan
            Uninstall-Module -Name $ModuleName -RequiredVersion $module.Version -Force
        }
        Write-Host "Old versions of '$ModuleName' have been removed. Latest version $($latestModule.Version) is retained." -ForegroundColor Green
    }
    catch {
        Write-Error "An error occurred while removing old versions: $_"
    }
}

function Install-UserModule {
    <#
    .SYNOPSIS
      Installs a module for the current user.
      
    .DESCRIPTION
      This wrapper function calls Install-Module with the -Scope CurrentUser parameter,
      ensuring that modules are installed for the current user.
      
    .PARAMETER Args
      Additional parameters for Install-Module.
      
    .EXAMPLE
      Install-UserModule -Name Pester -Force
      Installs the Pester module for the current user.
    #>
    [alias("ium")]
    param(
        [Parameter(ValueFromRemainingArguments = $true)]
        $Args
    )
    Install-Module -Scope CurrentUser @Args
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
