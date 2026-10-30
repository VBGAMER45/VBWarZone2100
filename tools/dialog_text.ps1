# Prints the texts of the dialog boxes (class #32770) of a running process and their controls: VB's
# run-time error box. powershell -File tools\dialog_text.ps1 Warzone
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class WinText {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc f, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    public static List<string> Dialogs(uint pid) {
        var r = new List<string>();
        EnumWindows((h, l) => {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != pid) return true;
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() != "#32770") return true;
            var t = new StringBuilder(1024); GetWindowText(h, t, 1024);
            r.Add("dialog: " + t);
            EnumChildWindows(h, (k, m) => {
                var s = new StringBuilder(4096); GetWindowText(k, s, 4096);
                if (s.Length > 0) r.Add("  " + s);
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        return r;
    }
}
"@
$name = "Warzone"
if ($args.Count -gt 0) { $name = $args[0] }
$p = Get-Process $name -ErrorAction SilentlyContinue
if ($null -eq $p) { "$name is not running"; exit 1 }
[WinText]::Dialogs([uint32]$p[0].Id)
