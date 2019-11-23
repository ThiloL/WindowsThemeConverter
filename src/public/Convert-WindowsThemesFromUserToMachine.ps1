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