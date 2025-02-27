# BlackBytesBox.Manifested.Initialize

A library for supporting CICD actions and PowerShell module management.

## PowerShell Module Utilities

A collection of PowerShell functions designed to simplify module and repository management:

- **Register-LocalGalleryRepository (`rlgr`)**  
  Registers a local PowerShell repository for gallery modules. Ensures the specified local repository folder exists, removes any existing repository with the given name, and registers the repository with a Trusted installation policy.

- **Update-ManifestModuleVersion (`ummv`)**  
  Updates the ModuleVersion in a PowerShell module manifest (psd1) file. Can work with either a direct file path or recursively search a directory for the first psd1 file.

- **Update-ModuleIfNewer (`umn`)**  
  Installs or updates a module from a repository (default: PSGallery) only if a newer version is available. Prevents unnecessary downloads when the installed module is already up to date.

- **Remove-OldModuleVersions (`romv`)**  
  Removes older versions of an installed PowerShell module, keeping only the latest version. Helps clean up local installations accumulated from repeated updates.

- **Install-UserModule (`ium`)**  
  A wrapper function for Install-Module that ensures modules are installed for the current user scope.

## Example Installation Command

```powershell
powershell -NoProfile -ExecutionPolicy Unrestricted -Command "& {
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force;
    Install-PackageProvider -Name NuGet -Force -MinimumVersion 2.8.5.201 -Scope CurrentUser | Out-Null;
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted;
    Install-Module PowerShellGet -Force -Scope CurrentUser -AllowClobber -WarningAction SilentlyContinue | Out-Null;
    Install-Module -Name BlackBytesBox.Manifested.Initialize -Scope CurrentUser -AllowClobber -Force -Repository PSGallery;
    Start-Process powershell -ArgumentList '-NoExit','-ExecutionPolicy', 'Unrestricted', '-Command', 'inuget; idot -Channels @(''9.0'') ; dotnet tool install --global BlackBytesBox.Distributed; satcom vscode'
}" ; exit
```

## Example Usage

```powershell
# Register a local gallery repository
Register-LocalGalleryRepository -RepositoryPath "$HOME/source/gallery" -RepositoryName "LocalGallery"

# Update a module manifest version
Update-ManifestModuleVersion -ManifestPath "C:\projects\MyModule\MyModule.psd1" -NewVersion "2.0.0"

# Update a module if a newer version exists
Update-ModuleIfNewer -ModuleName "MyModule"

# Clean up old versions of a module
Remove-OldModuleVersions -ModuleName "MyModule"

# Install a module for the current user
Install-UserModule -Name "MyModule" -Force
```
