# BlackBytesBox.Manifested
A library for supporting CICD actions

# PowerShell Module Utilities

A small collection of PowerShell functions designed to simplify module and repository management. This repo includes:

- **Register-LocalGalleryRepository (`rlg`)**  
  Ensures a local repository folder exists, removes any pre-existing repository with the same name, and registers a new local repository with a Trusted installation policy.

- **Convert-DateTimeToVersion64SecondsString (`cdv64`)**  
  Converts a UTC DateTime into version components (build, major, minor, revision) using a 64-second granularity for versioning suitable for NuGet packages and assemblies.

- **Update-ManifestModuleVersion (`ummv`)**  
  Recursively searches for a PSD1 manifest in a directory and updates its ModuleVersion while preserving comments and formatting.

- **Update-ModuleIfNewer (`umn`)**  
  Searches the specified repository (default: PSGallery) for a module update and installs it only if a newer version is available, avoiding unnecessary forced downloads.

- **Remove-OldModuleVersions**  
  Cleans up older installed versions of a module, retaining only the latest version.

## Example Commands

```powershell
powershell -NoProfile -ExecutionPolicy Unrestricted -Command "& {
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force;
    Install-PackageProvider -Name NuGet -Force -MinimumVersion 2.8.5.201 -Scope CurrentUser | Out-Null;
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted;
    Install-Module PowerShellGet -Force -Scope CurrentUser -AllowClobber -WarningAction SilentlyContinue | Out-Null;
    Install-Module -Name BlackBytesBox.Manifested -Scope CurrentUser -AllowClobber -Force -Repository PSGallery;
    Start-Process powershell -ArgumentList '-NoExit','-ExecutionPolicy', 'Unrestricted', '-Command', 'inuget; idot -Channels @(''9.0'') ; dotnet tool install --global BlackBytesBox.Distributed; satcom vscode'
}" ; exit
```

Register a local gallery repository:
```powershell

powershell -NoProfile -ExecutionPolicy unrestricted -Command "& Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force; Install-Module -Name STROM.NANO.PSWH.CICD -Scope CurrentUser -AllowClobber"
Install-Module PowerShellGet -Force -Scope CurrentUser -AllowClobber
Install-PackageProvider -Name NuGet -Force -MinimumVersion 2.8.5.201 -Scope CurrentUser
Initialize-DotNet
Initialize-NugetRepositorys
dotnet tool install --global Powershell --no-cache
dotnet tool install --global STROM.ATOM.TOOL.Common

Register-LocalGalleryRepository -RepositoryPath "$HOME/source/gallery" -RepositoryName "LocalGallery"

$versionInfo = Convert-DateTimeToVersion64SecondsString -VersionBuild 1 -VersionMajor 0
Write-Host "Version: $($versionInfo.VersionFull)"

Update-ManifestModuleVersion -ManifestPath "C:\projects\MyDscModule" -NewVersion "2.0.0"

Update-ModuleIfNewer -ModuleName "STROM.NANO.PSWH.CICD"

Remove-OldModuleVersions -ModuleName "STROM.NANO.PSWH.CICD"
```

### General BlackBytesBox naming conventions
---

BlackBytesBox.Manifested (PowerShell module)
BlackBytesBox.Constructed (MSBuild lib)
BlackBytesBox.Unified (NET Standard library)
BlackBytesBox.Distributed (Dotnet tool)
BlackBytesBox.Composed (NET library)
BlackBytesBox.Dosed (NET-Windows library)
BlackBytesBox.Routed (ASP.NET library)
BlackBytesBox.Bladed (ASP.NET Razor library)
BlackBytesBox.Retired (old .NET Framework 4.0 library)
BlackBytesBox.Seeded (template project)
BlackBytesBox.[Adjective].[Qualifier] (for further clarity when needed)

BlackBytesBox.Manifested.Base  (Powershell module)
BlackBytesBox.Distributed.Core  (Dotnet tool)