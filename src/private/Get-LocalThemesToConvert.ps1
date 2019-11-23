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
