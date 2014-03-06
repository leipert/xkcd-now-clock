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
[Wallpaper.Setter]::SetWallpaper( (Convert-Path $args[0]) )