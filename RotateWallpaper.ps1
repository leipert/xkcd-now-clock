# Taken from the Stack Exchange Network. Based on http://stackoverflow.com/a/9440226 by Andy Arismendi (https://stackoverflow.com/users/251123/andy-arismendi)
# Licensed under cc-by-sa 3.0 https://creativecommons.org/licenses/by-sa/3.0/
Add-Type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper
{
   public class Setter {
      public const int SetDesktopWallpaper = 20;
      public const int UpdateIniFile = 0x01;
      public const int SendWinIniChange = 0x02;
      [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
      private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
      public static void SetWallpaper ( string path ) {
         SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
         RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
         key.SetValue(@"WallpaperStyle", "0") ; 
         key.SetValue(@"TileWallpaper", "0") ; 
         key.Close();
      }
   }
}
"@

# Get root path of script
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition;
$path=$PSScriptRoot + "\images\";
write-output ("Image Path: " + $path);

# Get UTC Time + 12 hours
$utc = (get-date).ToUniversalTime().AddHours(12);

# Get elapsed seconds since last hour started
$elapsed = $utc.Minute * 60 + $utc.Second;

# choose right minute parameter based on elapsed seconds 
if ($elapsed -lt 450){ $m = "00";} # < 7.5 min
elseif ($elapsed -lt 1350){ $m = "15";} # < 22.5 min
elseif ($elapsed -lt 2250){ $m = "30";} # < 37.5 min
elseif ($elapsed -lt 3150){ $m = "45";} # < 52.5 min
else { $m = "00"; $utc = $utc.AddHours(1);} # > 52.5 min

# Create filename (for example 12h15m.jpg)
$file=($utc).ToString("HH") + "h" + $m + "m.jpg";

# Set Wallpaper
write-output ("Set wallpaper to: " + $file);
[Wallpaper.Setter]::SetWallpaper( (Convert-Path ($path + $file)) )