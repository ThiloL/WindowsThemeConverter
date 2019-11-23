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