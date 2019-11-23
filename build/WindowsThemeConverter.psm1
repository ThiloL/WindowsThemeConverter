function Get-LocalThemesToConvert {

  [CmdletBinding()]

  Param (
    # Your parameters go here...
  )

  # per-user theme folder
  $UsersThemesRootFolder = "$Env:LOCALAPPDATA\Microsoft\Windows\Themes"

  # getting the themes in
  $SubFolderNames = @(Get-ChildItem -Path $UsersThemesRootFolder -Directory -Force -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
  if (!$SubFolderNames) { return $null }

  # getting theme-filename
  $ThemeFileNames = @($SubFolderNames | ForEach-Object {
    Get-ChildItem -Path "$Env:LOCALAPPDATA\Microsoft\Windows\Themes\$_" -Filter *.theme | Select-Object -ExpandProperty FullName
  })
  if (!$ThemeFileNames) { return $null }

  # builing object
  $Themes = $ThemeFileNames | ForEach-Object {
    $Ini = Get-IniContent -FilePath $_ -Verbose:$false
    New-Object -TypeName psobject -Property @{Name = $Ini['Theme']['DisplayName']; FileName = $_ }
  }

  # selecting themes
  $SelectedThemes = @($Themes | Out-GridView -Title "Select the Theme(s) to convert" -PassThru)
  if (!$SelectedThemes) { return $null }

  # return it
  return $SelectedThemes
}

Function Invoke-FixAppEvent {
  [CmdletBinding()]

  param(
    $Ini,
    [string]$ThemeFilename,
    $Guid
  )

  Write-Verbose -ForegroundColor gray "Fixing AppEvents of '$ThemeFilename'..."

  [System.Collections.ArrayList]$RelevantKeys = @()

  # finding AppEvents
  foreach ($key in $Ini.Keys) {
    if (([string]$key).StartsWith("AppEvents")) {
      $RelevantKeys.Add($key)
    }
  }

  # fixing these entries, with the correct path
  $RelevantKeys | ForEach-Object {
    if ($Ini["$_"]["DefaultValue"]) {
      $Value = $Ini["$_"]["DefaultValue"]
      $AppEventsFileName = Split-Path -Path ([System.Environment]::ExpandEnvironmentVariables($Value)) -Leaf
      $Ini["$_"]["DefaultValue"] = "%systemroot%\resources\themes\$Guid\$AppEventsFileName"
    }
  }
}

# Private Function Example - Replace With Your Function
function Invoke-PatchTheme {

  [CmdletBinding()]

  Param (
    [string]$ThemeFilename
  )

  # new GUID for this theme
  $Guid = (New-Guid).Guid

  Write-Verbose -ForegroundColor Gray "Patching '$ThemeFilename' with GUID '$Guid'..."

  # reading theme-file
  $Ini = Get-IniContent -FilePath $ThemeFilename -Verbose:$false

  if ($Ini['Slideshow']) {
      if ($Ini['Slideshow']['ImagesRootPIDL']) {
          $Ini['Slideshow']['ImagesRootPIDL'] = $null
      }
      else {
          Write-Warning "'ImagesRootPIDL' not found in section 'Slideshow'!"
          return 1
      }
  }
  else {
      Write-Warning "Section 'Slideshiw' not found!"
      return 1
  }

  $Ini['Slideshow']['ImagesRootPath'] = "%systemroot%\resources\themes\$Guid\DesktopBackground"

  if ($Ini['Slideshow.A']) {
      if ($Ini['Slideshow.A']['ImagesRootPath']) {
          $Ini['Slideshow.A']['ImagesRootPath'] = "%systemroot%\resources\themes\$Guid\DesktopBackground"
      }
  }

  if ($Ini['Slideshow.W']) {
      if ($Ini['Slideshow.W']['ImagesRootPath']) {
          $Ini['Slideshow.W']['ImagesRootPath'] = "%systemroot%\resources\themes\$Guid\DesktopBackground"
      }
  }

  $ThemeSourceFolder = Split-Path -Path $ThemeFilename -Parent
  $FirstBackgroundImage = (Get-ChildItem -Path (Join-Path -Path $ThemeSourceFolder -ChildPath "DesktopBackground") -Filter "*.jpg" | Select-Object -First 1).Name

  if ($Ini['Control Panel\Desktop']) {
      if ($Ini['Control Panel\Desktop']['Wallpaper']) {
          $Ini['Control Panel\Desktop']['Wallpaper'] = "%systemroot%\resources\themes\$Guid\DesktopBackground\$FirstBackgroundImage"
      }
  }

  if ($Ini['Control Panel\Desktop.A']) {
      if ($Ini['Control Panel\Desktop.A']['Wallpaper']) {
          $Ini['Control Panel\Desktop.A']['Wallpaper'] = "%systemroot%\resources\themes\$Guid\DesktopBackground\$FirstBackgroundImage"
      }
  }

  if ($Ini['Control Panel\Desktop.W']) {
      if ($Ini['Control Panel\Desktop.W']['Wallpaper']) {
          $Ini['Control Panel\Desktop.W']['Wallpaper'] = "%systemroot%\resources\themes\$Guid\DesktopBackground\$FirstBackgroundImage"
      }
  }

  # Handling AppEvents (WAV)

  $AppEvents = @($Ini.Keys -match "AppEvents")

  if ($AppEvents) {

      Write-Verbose -ForegroundColor Yellow "AppEvents found. Fixing it..."
      Invoke-FixAppEvent -Ini $Ini -ThemeFilename $ThemeFilename -Guid $Guid
  }

  # new Theme filename
  $NewThemeFilename = Join-Path -Path $TempFolder -ChildPath "$Guid.theme"

  # writing new Theme file
  $Ini | Out-IniFile -FilePath $NewThemeFilename -Force -Verbose:$false

  # new Theme sub folder
  $NewThemeFolder = Join-Path $TempFolder -ChildPath $Guid

  # creating new Theme sub folder (images, sounds, ...)
  New-Item -Force -ItemType Directory -Path $NewThemeFolder | Out-Null

  Write-Verbose "Copying stuff from '$ThemeSourceFolder' to '$NewThemeFolder'..."

  Copy-Item -Path "$ThemeSourceFolder\*" -Destination "$NewThemeFolder" -Force -Recurse

  # removing theme files
  Remove-Item -Path $NewThemeFolder -Include "*.theme" -Force -Recurse

  return 0

}

# Public Function Example - Replace With Your Function
function Convert-WindowsThemesFromUserToMachine {

  [CmdletBinding()]

  Param (
    # Your parameters go here...
  )

  # getting the thems to convert
  $ThemesToConvert = Get-LocalThemesToConvert

  if (!$ThemesToConvert)
  {
    write-Error "No themes found or selected for convertion!"
    Exit 1
  }

  # temp folder for output
  $TempFolder = Join-Path -Path $env:TEMP -ChildPath (New-Guid).Guid

  Write-Verbose -ForegroundColor Gray "Creating '$TempFolder' as global output folder..."
  New-Item -ItemType Directory $TempFolder | Out-Null

  $ThemesToConvert | ForEach-Object {

    $Return = Invoke-PatchTheme -ThemeFilename $_.FileName

    if ($Return -eq 0)
    {
      Write-Output -ForegroundColor Green "'$($_.Name)' succesfully patched."
    }
    else
    {
      Write-Output -ForegroundColor Red "'$($_.Name)' NOT succesfully patched."
    }
  }

  Write-Output -ForegroundColor Green "Please copy the contents of '$TempFolder' to 'C:\Windows\Resources\Themes' to install the Themes per-machine!"

  Start-Process -FilePath "$env:windir\explorer.exe" -ArgumentList "$TempFolder"
}

Export-ModuleMember -Function Convert-WindowsThemesFromUserToMachine
