# linkwatch.ps1 - records the command line of every LINK.EXE (VB6's linker) that starts while it watches: one file
# per process, build\linkwatch\link_<pid>.txt, and a line in build\linkwatch\links.log (its length and its .OBJ
# count). VB6 passes every module's .OBJ to LINK.EXE on that one line, which CreateProcess caps at 32,767
# characters (docs/PLAN.md, Risks). Start it, build (tools/build.py) meanwhile; it stops after -hours.
# A WMI creation event (WMI polls every 0.1 s; no administrator rights needed) hands over the new process with
# its command line: a loop of Get-CimInstance calls missed the shorter links.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\linkwatch.ps1 [-hours 1]
param([double]$hours = 1)
$out = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) "build\linkwatch"
New-Item -ItemType Directory -Force $out | Out-Null
$query = "SELECT * FROM __InstanceCreationEvent WITHIN 0.1 WHERE TargetInstance ISA 'Win32_Process' " +
         "AND TargetInstance.Name = 'LINK.EXE'"
Register-CimIndicationEvent -Query $query -SourceIdentifier LinkWatch | Out-Null
$end = (Get-Date).AddHours($hours)
try {
    while ((Get-Date) -lt $end) {
        $e = Wait-Event -SourceIdentifier LinkWatch -Timeout 5
        if ($e) {
            $p = $e.SourceEventArgs.NewEvent.TargetInstance
            $cl = $p.CommandLine
            if ($cl) {
                [IO.File]::WriteAllText((Join-Path $out ("link_" + $p.ProcessId + ".txt")), $cl)
                $objs = ([regex]::Matches($cl, '\.OBJ"', 'IgnoreCase')).Count
                $line = "{0:yyyy-MM-dd HH:mm:ss} pid {1} chars {2} objs {3}" -f (Get-Date), $p.ProcessId, $cl.Length, $objs
                Add-Content -Path (Join-Path $out "links.log") -Value $line
            }
            Remove-Event -EventIdentifier $e.EventIdentifier
        }
    }
}
finally {
    Unregister-Event -SourceIdentifier LinkWatch
}
