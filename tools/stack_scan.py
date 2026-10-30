"""Lists the return addresses into the game's modules (Warzone.exe and the Wz*.dll of a debug build) found
on the main thread's stack of the running Warzone.exe, innermost first: a call chain with some noise
(stale frames), for when cdb can't unwind VB's frames (a tools\exe_debug.py build: the .pdb).
python -I tools\stack_scan.py [bytes]"""
import re
import subprocess
import sys

CDB = r"C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\cdb.exe"
BIN = r"C:\Users\Owner\desktop\websites\warzone\bin"
size = sys.argv[1] if len(sys.argv) > 1 else "0x8000"
out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq Warzone.exe", "/FO", "CSV", "/NH"], capture_output=True, text=True).stdout
pid = None
for line in out.splitlines():
    parts = [p.strip('"') for p in line.split('","')]
    if len(parts) > 1 and parts[0].lower() == "warzone.exe":
        pid = parts[1]
if pid is None:
    print("Warzone.exe is not running")
    sys.exit(1)
r = subprocess.run([CDB, "-pv", "-p", pid, "-y", BIN, "-c", ".reload /f; ~0 s; dps @esp L%s/4; qd" % size],
                   capture_output=True, text=True, timeout=300)
last = None
for l in r.stdout.splitlines():
    m = re.match(r"^([0-9a-f]{8})\s+([0-9a-f]{8})\s+(\S+)", l)
    if m and re.match(r"(Warzone|Wz\w+)!", m.group(3), re.I):
        if m.group(3) != last:
            print(m.group(1), m.group(3))
        last = m.group(3)
