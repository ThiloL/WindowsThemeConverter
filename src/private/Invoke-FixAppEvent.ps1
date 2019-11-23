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