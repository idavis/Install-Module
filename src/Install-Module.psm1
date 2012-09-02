#Requires -Version 2.0

function Install-Module {
  param(
    [Parameter(Position=0,Mandatory=$true,ParameterSetName='file')][string]$file,
    [Parameter(Position=0,Mandatory=$true,ParameterSetName='directory')][string]$directory,
	[Parameter(Position=1,Mandatory=$false,ParameterSetName='directory')][switch]$recurse,
	[Parameter(Position=2,Mandatory=$false)][ValidateSet('User','Global','All')][string]$scope = 'User'
  )
  if($scope -eq "All" -or $scope -eq "Global"){
    if(!(Test-IsAdmin)) {
      Write-Error "Installing module $name in $scope scope requires elevated privileges. Installation aborted."
      return
    }
  }

  if($PsCmdlet.ParameterSetName -eq "directory") {
    # infer name from the parent directory since it wasn't specified.
    $path = New-Object System.IO.DirectoryInfo(Resolve-Path $directory)
    Assert-PathExists $path
    [string]$name = $path.BaseName
    Assert-ModuleNamesMatch $name $path
  } else {
    $fileInfo = New-Object System.IO.FileInfo(Resolve-Path $file)
    if(!$fileInfo.Exists) {
      throw "The module file $fileInfo was not found. Installation aborted."
    }
    [string]$name = [IO.Path]::GetFileNameWithoutExtension($fileInfo)
  }

  switch($PsCmdlet.ParameterSetName) {
    "file" {
      switch($scope) {
        'User' { Install-ModuleForCurrentUser $name $fileInfo }
	    'Global' { Install-ModuleForAllUsers $name $fileInfo }
	    'All' {
	      Install-ModuleForCurrentUser $name $fileInfo
	      Install-ModuleForAllUsers $name $fileInfo
	    }
      }
    }
    "directory" {
      switch($scope) {
        'User' { Install-ModuleForCurrentUser $name $path $recurse }
	    'Global' { Install-ModuleForAllUsers $name $path $recurse }
	    'All' {
	      Install-ModuleForCurrentUser $name $path $recurse
	      Install-ModuleForAllUsers $name $path $recurse
	    }
      }
    }
  }
}

function Uninstall-Module {
  param(
    [Parameter(Mandatory=$true)][string]$name,
	[Parameter(Mandatory=$false)][ValidateSet('User','Global','All')][string]$scope = 'User'
  )
  switch($scope) {
    'User' { Uninstall-ModuleForCurrentUser $name }
	'Global' { Uninstall-ModuleForAllUsers $name }
	'All' {
	  Uninstall-ModuleForCurrentUser $name
	  Uninstall-ModuleForAllUsers $name
	}
  }
}

function Install-ModuleForCurrentUser {
  param(
    [Parameter(Position=0,Mandatory=$true)][string]$name,
    [Parameter(Position=1,Mandatory=$true,ParameterSetName='file')][IO.FileInfo]$file,
    [Parameter(Position=1,Mandatory=$true,ParameterSetName='directory')][IO.DirectoryInfo]$path,
	[Parameter(Mandatory=$false,ParameterSetName='directory')][switch]$recurse
  )
  $modulesDirectory = Get-ModulesDirectoryForCurrentUser
  switch($PsCmdlet.ParameterSetName) {
    "file" {Install-ModuleToPath $name $file $modulesDirectory}
    "directory" {Install-ModuleToPath $name $path $recurse $modulesDirectory}
  }
}

function Install-ModuleForAllUsers {
  param(
    [Parameter(Position=0,Mandatory=$true)][string]$name,
    [Parameter(Position=1,Mandatory=$true,ParameterSetName='file')][IO.FileInfo]$file,
    [Parameter(Position=1,Mandatory=$true,ParameterSetName='directory')][IO.DirectoryInfo]$path,
	[Parameter(Mandatory=$false,ParameterSetName='directory')][switch]$recurse
  )
  $modulesDirectory = Get-ModulesDirectoryForAllUsers $name
  switch($PsCmdlet.ParameterSetName) {
    "file" {Install-ModuleToPath  $name $file $modulesDirectory}
    "directory" {Install-ModuleToPath $name $path $recurse $modulesDirectory}
  }
}

function Uninstall-ModuleForCurrentUser {
  param(
    [Parameter(Mandatory=$true)][string]$name
  )
  $modulesDirectory = Get-ModulesDirectoryForCurrentUser
  Remove-ItemRecurseForce (Join-Path $modulesDirectory $name)
}

function Uninstall-ModuleForAllUsers {
  param(
    [Parameter(Mandatory=$true)][string]$name
  )
  $modulesDirectory = Get-ModuleRootDirectoryForAllUsers $name
  Remove-ItemRecurseForce $modulesDirectory
  Remove-PathFromPSModulePath $modulesDirectory
}

function Install-ModuleToPath {
  param(
    [Parameter(Position=0,Mandatory=$true)][string]$name,
    [Parameter(Position=1,Mandatory=$true,ParameterSetName='file')][IO.FileInfo]$file,
    [Parameter(Position=1,Mandatory=$true,ParameterSetName='directory')][IO.DirectoryInfo]$path,	
    [Parameter(Position=2,Mandatory=$false,ParameterSetName='directory')][switch]$recurse,
    [Parameter(Position=3,Mandatory=$true)][string]$modulesPath
  )
  $targetDirectory = Join-Path $modulesDirectory $name
  Create-DirectoryIfNeeded $modulesDirectory
  Add-ModulesDirectoryToPSModulePath $modulesDirectory
  Create-DirectoryIfNeeded $targetDirectory
  switch($PsCmdlet.ParameterSetName) {
    "file" {Copy-Module $file $targetDirectory}
    "directory" {Copy-Module $path $recurse $targetDirectory }
  }
}

function Copy-Module {
  param(
    [Parameter(Position=0,Mandatory=$true,ParameterSetName='file')][IO.FileInfo]$file,
    [Parameter(Position=0,Mandatory=$true,ParameterSetName='directory')][IO.DirectoryInfo]$path,	
    [Parameter(Position=1,Mandatory=$false,ParameterSetName='directory')][switch]$recurse,
    [Parameter(Position=2,Mandatory=$true)][IO.DirectoryInfo]$destination
  )
  switch($PsCmdlet.ParameterSetName) {
    "file" {
      $targetFileName = Join-Path $destination $file.Name
      $file.CopyTo($targetFileName,$true)
    }
    "directory" {
      if($recurse) {
        Copy-Item $source\* $destination -Recurse -Force
      } else {
        Copy-Item $source\* $destination -Force
      }
    }
  }
}

function Remove-ItemRecurseForce {
  param($path)
  Remove-Item -Recurse -Force "$path"
}

function Add-ModulesDirectoryToPSModulePath {
  param(
    [Parameter(Mandatory=$true)][string]$path
  )
  $isInPath = ($env:PSModulePath.Split(";") | ? { $_.Trim('\') -eq $path } | Measure-Object).Count -gt 0
  if(!$isInPath) {
    $env:PSModulePath += ";$path"
  }
}

function Remove-PathFromPSModulePath {
  param(
    [Parameter(Mandatory=$true)][string]$path
  )
  $isInPath = $env:PSModulePath.ToUpperInvariant().Contains($path.Trim('\').ToUpperInvariant())
  if($isInPath) {
    $env:PSModulePath = ($env:PSModulePath.Split(";") | ? { !$_.ToUpperInvariant().Contains($path.ToUpperInvariant().Trim('\')) }) -join ';'
  }
}

function Get-ModulesDirectoryForCurrentUser {
  $modulesDirectory = Join-Path ([Environment]::GetFolderPath("MyDocuments")) WindowsPowerShell\Modules
  $modulesDirectory
}

function Get-ModuleRootDirectoryForAllUsers {
  param($name)
  # We are not going to use ${env:SystemRoot}\WindowsPowerShell\1.0\Modules 
  # as it is for Microsoft modules, and we should not mess with system Directories
  $moduleRoot = Join-Path $env:ProgramFiles $name
  return $moduleRoot
}

function Get-ModulesDirectoryForAllUsers {
  param($name)
  $moduleRoot = Get-ModuleRootDirectoryForAllUsers $name
  $modulesDirectory = Join-path $moduleRoot "PowerShell\Modules"
  $modulesDirectory
}

function Create-DirectoryIfNeeded {
  param($path)
  if(!(Test-Path $path)) {
    New-Item -path $path -ItemType Directory -Force
  }
}

function Assert-PathExists {
  param($path)
  if(!(Test-Path $path)) {
    throw ("The path does not exist: $path.")
  }
}

function Assert-ModuleNamesMatch {
  param(
    [Parameter(Mandatory=$true)][string]$moduleName,
	[Parameter(Mandatory=$true)][string]$modulePath
  )
  $modules = Get-ChildItem -Path $modulePath -Include *.psm1
  $hasMatchingModule = ($modules | ? { $_.Name -eq $moduleName} | Measure-Object).Count -gt 0
  if($hasMatchingModule) { return }
  $manifests = Get-ChildItem -Path $modulePath -Include *.psd1
  $hasMatchingManifest = ($manifests | ? { $_.Name -eq $moduleName} | Measure-Object).Count -gt 0
  if($hasMatchingManifest) { return }
  $dlls = Get-ChildItem -Path $modulePath -Include *.dll
  $hasMatchingLibrary = ($dlls | ? { $_.Name -eq $moduleName} | Measure-Object).Count -gt 0
  if($hasMatchingLibrary) { return }
  throw ("The module to be installed must have a file named $moduleName (.psm1, psd1 or, .dll).")
}

function Test-IsAdmin {
  $windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  $windowsPrincipal = New-Object 'System.Security.Principal.WindowsPrincipal' $windowsIdentity
  return $windowsPrincipal.IsInRole("Administrators")
}

Export-ModuleMember -Function Install-Module, Uninstall-Module