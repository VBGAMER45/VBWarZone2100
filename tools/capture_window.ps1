# Starts Warzone.exe with the given arguments, captures its window after a
# delay and waits for it to quit (kills it after 30 s). Checks the real window
# in scripted runs; the arguments should end the run (--frames=N).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\capture_window.ps1 -exe bin\Warzone.exe -arguments "--frames=300" -delay 2000 -out shot.png
param([string]$exe, [string]$arguments, [int]$delay = 2000, [string]$out)
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public class W { [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string c, IntPtr t);
 public struct RECT { public int L, T, R, B; } }
"@
$exe = (Resolve-Path $exe).Path
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$p = Start-Process -FilePath $exe -ArgumentList $arguments -PassThru -WorkingDirectory (Split-Path $exe)
Start-Sleep -Milliseconds $delay
$h = [W]::FindWindow("WarzoneVB", [IntPtr]::Zero)   # a $null string would arrive as ""
if ($h -eq [IntPtr]::Zero) { "no WarzoneVB window found" }
else {
  [W]::SetForegroundWindow($h) | Out-Null
  Start-Sleep -Milliseconds 300
  $r = New-Object W+RECT
  [W]::GetWindowRect($h, [ref]$r) | Out-Null
  $bmp = New-Object System.Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($r.L, $r.T, 0, 0, $bmp.Size)
  $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
  $c = $bmp.GetPixel([int](($r.R - $r.L) / 2), [int](($r.B - $r.T) / 2))
  "window $($r.R - $r.L) x $($r.B - $r.T), centre pixel R=$($c.R) G=$($c.G) B=$($c.B)"
}
if (-not $p.WaitForExit(30000)) { "still running: killing"; $p.Kill() } else { "exited with code $($p.ExitCode) after $([int]$sw.ElapsedMilliseconds) ms" }
