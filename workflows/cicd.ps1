param (
    [string]$NUGET_GITHUB_PUSH,
    [string]$NUGET_PAT,
    [string]$NUGET_TEST_PAT,
    [string]$POWERSHELL_GALLERY
)

# If any of the parameters are empty, try loading them from a secrets file.
if ([string]::IsNullOrEmpty($NUGET_GITHUB_PUSH) -or [string]::IsNullOrEmpty($NUGET_PAT) -or [string]::IsNullOrEmpty($NUGET_TEST_PAT) -or [string]::IsNullOrEmpty($POWERSHELL_GALLERY)) {
    if (Test-Path "$PSScriptRoot\cicd_secrets.ps1") {
        . "$PSScriptRoot\cicd_secrets.ps1"
        Write-Host "Secrets loaded from file."
    }
    if ([string]::IsNullOrEmpty($NUGET_GITHUB_PUSH))
    {
        exit 1
    }
}

Install-Module -Name BlackBytesBox.Manifested.Initialize -Repository "PSGallery" -Force -AllowClobber
Install-Module -Name BlackBytesBox.Manifested.Version -Repository "PSGallery" -Force -AllowClobber
Install-Module -Name BlackBytesBox.Manifested.Git -Repository "PSGallery" -Force -AllowClobber


$result1 = Convert-DateTimeTo64SecPowershellVersion -VersionBuild 0
$result2 = Get-GitCurrentBranch
$result4 = Get-GitCurrentBranchRoot
$result3 = Get-GitTopLevelDirectory


##############################

# Define the path to your module folder (adjust "MyModule" as needed)
$moduleFolder = "$result3/source/BlackBytesBox.Manifested.Git"
Update-ManifestModuleVersion -ManifestPath "$moduleFolder" -NewVersion "$($result1.VersionBuild).$($result1.VersionMajor).$($result1.VersionMinor)"
$moduleManifest = "$moduleFolder/BlackBytesBox.Manifested.Git.psd1" -replace '[/\\]', [System.IO.Path]::DirectorySeparatorChar

# Validate the module manifest
Write-Host "===> Testing module manifest at: $moduleManifest" -ForegroundColor Cyan
Test-ModuleManifest -Path $moduleManifest

Publish-Module -Path $moduleFolder -Repository "PSGallery" -NuGetApiKey "$POWERSHELL_GALLERY"

exit 0

# Use for cleaning local enviroment only, use channelRoot for deployment.
$isCiCd = $false
$isLocal = $false
if ($env:GITHUB_ACTIONS -ieq "true")
{
    $isCiCd = $true
}
else {
    $isLocal = $true
}

# Check if the secrets file exists before importing
if (Test-Path "$PSScriptRoot/cicd_secrets.ps1") {
    . "$PSScriptRoot\cicd_secrets.ps1"
    Write-Host "Secrets loaded from file."
} else {
    $NUGET_GITHUB_PUSH = $args[0]
    $NUGET_PAT = $args[1]
    $NUGET_TEST_PAT = $args[2]
    Write-Host "Secrets will be taken from args."
}

$currentBranch = Get-GitCurrentBranch

$currentBranchRoot = Get-BranchRoot -BranchName "$currentBranch"
$topLevelDirectory = Get-GitTopLevelDirectory

#Branch too channel mappings
$branchSegments = @(Split-Segments -InputString "$currentBranch" -ForbiddenSegments @("latest") -MaxSegments 2)
$nugetSuffix = @(Translate-FirstSegment -Segments $branchSegments -TranslationTable @{ "feature" = "-development"; "develop" = "-quality"; "bugfix" = "-quality"; "release" = "-staging"; "main" = ""; "master" = ""; "hotfix" = "" } -DefaultTranslation "{nodeploy}")
$nugetSuffix = $nugetSuffix[0]
$channelSegments = @(Translate-FirstSegment -Segments $branchSegments -TranslationTable @{ "feature" = "development"; "develop" = "quality"; "bugfix" = "quality"; "release" = "staging"; "main" = "production"; "master" = "production"; "hotfix" = "production" } -DefaultTranslation "{nodeploy}")

$branchFolder = Join-Segments -Segments $branchSegments
$branchVersionFolder = Join-Segments -Segments $branchSegments -AppendSegments @( $calculatedVersion.VersionFull )
$channelRoot = $channelSegments[0]
$channelVersionFolder = Join-Segments -Segments $channelSegments -AppendSegments @( $calculatedVersion.VersionFull )
$channelVersionFolderRoot = Join-Segments -Segments $channelSegments -AppendSegments @( "latest" )
if ($channelSegments.Count -eq 2)
{
    $channelVersionFolderRoot = Join-Segments -Segments $channelRoot -AppendSegments @( "latest" )
}


Write-Output "BranchFolder to $branchFolder"
Write-Output "BranchVersionFolder to $branchVersionFolder"
Write-Output "ChannelRoot to $channelRoot"
Write-Output "ChannelVersionFolder to $channelVersionFolder"
Write-Output "ChannelVersionFolderRoot to $channelVersionFolderRoot"

#Guard for variables
Ensure-Variable -Variable { $calculatedVersion } -ExitIfNullOrEmpty
Ensure-Variable -Variable { $currentBranch } -ExitIfNullOrEmpty
Ensure-Variable -Variable { $currentBranchRoot } -ExitIfNullOrEmpty
Ensure-Variable -Variable { $topLevelDirectory } -ExitIfNullOrEmpty
Ensure-Variable -Variable { $nugetSuffix }
Ensure-Variable -Variable { $NUGET_GITHUB_PUSH } -ExitIfNullOrEmpty -HideValue
Ensure-Variable -Variable { $NUGET_PAT } -ExitIfNullOrEmpty -HideValue
Ensure-Variable -Variable { $NUGET_TEST_PAT } -ExitIfNullOrEmpty -HideValue

#Required directorys
$artifactsOutputFolderName = "artifacts"
$reportsOutputFolderName = "reports"

$outputRootArtifactsDirectory = [System.IO.Path]::Combine($topLevelDirectory, $artifactsOutputFolderName)
$outputRootReportResultsDirectory   = [System.IO.Path]::Combine($topLevelDirectory, $reportsOutputFolderName)
$targetConfigAllowedLicenses = [System.IO.Path]::Combine($topLevelDirectory, ".config", "allowed-licenses.json")

Ensure-Variable -Variable { $outputRootArtifactsDirectory } -ExitIfNullOrEmpty
Ensure-Variable -Variable { $outputRootReportResultsDirectory } -ExitIfNullOrEmpty
Ensure-Variable -Variable { $targetConfigAllowedLicenses } -ExitIfNullOrEmpty

[System.IO.Directory]::CreateDirectory($outputRootArtifactsDirectory) | Out-Null
[System.IO.Directory]::CreateDirectory($outputRootReportResultsDirectory) | Out-Null

if (-not $isCiCd) { Delete-FilesByPattern -Path "$outputRootArtifactsDirectory" -Pattern "*"  }
if (-not $isCiCd) { Delete-FilesByPattern -Path "$outputRootReportResultsDirectory" -Pattern "*"  }

