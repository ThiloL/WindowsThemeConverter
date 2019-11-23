Function Fix-AppEvents
{
    param(
        $Ini,
        [string]$ThemeFilename
    )

    Write-Host -ForegroundColor gray "Fixing AppEvents of '$ThemeFilename'..."

    [System.Collections.ArrayList]$RelevantKeys = @()

    foreach($key in $Ini.Keys)
    {
        if (([string]$key).StartsWith("AppEvents"))
        {
            $RelevantKeys.Add($key)
        }
    }

    $RelevantKeys | % {

        if ($Ini["$_"]["DefaultValue"])
        {
            Write-Host "yeah"
            $Value = $Ini["$_"]["DefaultValue"]
            $AppEventsFileName = Split-Path -Path ([System.Environment]::ExpandEnvironmentVariables($Value)) -Leaf
             $Ini["$_"]["DefaultValue"] = "%systemroot%\resources\themes\$Guid\$AppEventsFileName"
        }
    }
}

Function Patch-Theme {
    param([string]$ThemeFilename)

    $Guid = (New-Guid).Guid

    Write-Host -ForegroundColor Gray "Patching '$ThemeFilename'..."
    Write-Host -ForegroundColor Gray "Using GUID '$Guid' for this theme"

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

    $AppEvents = @($Ini.Keys -match "AppEvents")

    if ($AppEvents) {

        Write-Warning "AppEvents found!"
        Fix-AppEvents -Ini $Ini -ThemeFilename $ThemeFilename
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

# === MAIN ===

Import-Module PsIni

$UsersThemesRootFolder = "$Env:LOCALAPPDATA\Microsoft\Windows\Themes"
$SubFolderNames = @(Get-ChildItem -Path $UsersThemesRootFolder -Directory -Force -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)

$ThemeFileNames = @($SubFolderNames | % { 

        Get-ChildItem -Path "$Env:LOCALAPPDATA\Microsoft\Windows\Themes\$_" -Filter *.theme | Select-Object -ExpandProperty FullName
    })

$Themes = $ThemeFileNames | % {

    $Ini = Get-IniContent -FilePath $_ -Verbose:$false
    New-Object -TypeName psobject -Property @{Name = $Ini['Theme']['DisplayName']; FileName = $_ }
}

$SelectedTheme = @($Themes | Out-GridView -Title "Select the Theme to convert" -PassThru)

if (-not($SelectedTheme)) { Exit }

$TempFolder = Join-Path -Path $env:TEMP -ChildPath (New-Guid).Guid

Write-Host -ForegroundColor Gray "Creating '$TempFolder' as global output folder..."
New-Item -ItemType Directory $TempFolder | Out-Null

$SelectedTheme | % {

    $R = Patch-Theme -ThemeFilename $_.FileName
}

Start-Process -FilePath "$env:windir\explorer.exe" -ArgumentList "$TempFolder"

Write-Host -ForegroundColor Green "Please copy the contents of '$TempFolder' to 'C:\Windows\Resources\Themes' to install the Themes per-machine!"