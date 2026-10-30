# Starts Warzone.exe with the given arguments, resizes its window after a
# delay and waits for it to quit (kills it after 30 s). Checks that a resized
# window persists its size; the arguments should end the run (--frames=N).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\resize_window.ps1 -exe bin\Warzone.exe -arguments "--frames=240" -width 900 -height 700
param([string]$exe, [string]$arguments, [int]$delay = 1500, [int]$width = 900, [int]$height = 700)
Add-Type @"
using System; using System.Runtime.InteropServices;
public class W { [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
 [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string c, IntPtr t);
 public struct RECT { public int L, T, R, B; } }
"@
$exe = (Resolve-Path $exe).Path
$p = Start-Process -FilePath $exe -ArgumentList $arguments -PassThru -WorkingDirectory (Split-Path $exe)
Start-Sleep -Milliseconds $delay
$h = [W]::FindWindow("WarzoneVB", [IntPtr]::Zero)   # a $null string would arrive as ""
if ($h -eq [IntPtr]::Zero) { "no WarzoneVB window found" }
else {
  # SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE
  [W]::SetWindowPos($h, [IntPtr]::Zero, 0, 0, $width, $height, 0x0002 -bor 0x0004 -bor 0x0010) | Out-Null
  Start-Sleep -Milliseconds 200
  $r = New-Object W+RECT
  [W]::GetClientRect($h, [ref]$r) | Out-Null
  "window resized to $width x $height, client area $($r.R - $r.L) x $($r.B - $r.T)"
}
if (-not $p.WaitForExit(30000)) { $p.Kill(); "killed after 30 s" } else { "exit code $($p.ExitCode)" }
