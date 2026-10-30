"""Cross-project probe: a VB6 Standard EXE using a VB6 ActiveX DLL.

The port is split into Warzone.exe and WzPlatform.dll (VB6's name table holds
about 32,100 names per project). This builds a probe DLL and EXEs that use it,
and answers:
  1. cost of calling a GlobalMultiUse function of the DLL from the exe
  2. a UDT array owned by the DLL, seen from the exe through a SAFEARRAY alias
  3. a callback from the DLL into the exe (an interface the exe implements)
  4. an error raised in the DLL, caught in the exe
  5. whether the DLL's names count in the exe's name table: exes with about
     29,000 own names that call 4,000 functions of the DLL, or none of them
  6. reg-free COM: the exe runs with a manifest after the DLL is unregistered

usage: python -I xproj.py <work dir>
"""
import os
import re
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import names_gen  # noqa: E402

VB6 = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
STDOLE = r"Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation"
OPTIONS = """CompilationType=0
OptimizationType=2
BoundsCheck=0
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
"""
API_CLASSES = 4
API_PER_CLASS = 1000


def cls_header(name, instancing):
    """VB6 class file header; instancing: private, notcreatable, global."""
    creatable, exposed, globalns = {"private": ("True", "False", "False"),
                                    "notcreatable": ("False", "True", "False"),
                                    "global": ("True", "True", "True")}[instancing]
    return ["VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "  Persistable = 0  'NotPersistable",
            "  DataBindingBehavior = 0  'vbNone", "  DataSourceBehavior  = 0  'vbNone",
            "  MTSTransactionMode  = 0  'NotAnMTSObject", "END",
            'Attribute VB_Name = "%s"' % name, "Attribute VB_GlobalNameSpace = %s" % globalns,
            "Attribute VB_Creatable = %s" % creatable, "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = %s" % exposed, "Option Explicit", ""]


def write(folder, name, lines):
    with open(os.path.join(folder, name), "w", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")


TILE_TYPE = ["Public Type TILE", "    texture As Long", "    height As Long", "    info As Byte",
             "    tileVisBits As Byte", "    pad As Integer", "    limitedContinent As Long", "End Type"]


def make_dll(folder):
    write(folder, "pmod.bas", ['Attribute VB_Name = "pmod"', "Option Explicit"] + TILE_TYPE +
          ["Public tiles() As TILE", "Public tileCount As Long", "Public sink As ISink"])
    write(folder, "ISink.cls", cls_header("ISink", "notcreatable") +
          ["Public Sub OnEvent(ByVal i As Long)", "End Sub"])
    write(folder, "Platform.cls", cls_header("Platform", "global") + [
        "Public Function PlatAdd(ByVal a As Long, ByVal b As Long) As Long", "    PlatAdd = a + b", "End Function", "",
        "Public Function MapCreate(ByVal n As Long) As Long", "    Dim i As Long", "    ReDim tiles(0 To n - 1)",
        "    tileCount = n", "    For i = 0 To n - 1", "        tiles(i).texture = i", "        tiles(i).height = i * 2",
        "        tiles(i).info = i And 255", "    Next", "    MapCreate = VarPtr(tiles(0))", "End Function", "",
        "Public Sub MapTouch()", "    Dim i As Long", "    For i = 0 To tileCount - 1",
        "        tiles(i).height = tiles(i).height + 1", "    Next", "End Sub", "",
        "Public Function TileSize() As Long", "    TileSize = LenB(tiles(0))", "End Function", "",
        "Public Sub SetSink(ByVal s As ISink)", "    Set sink = s", "End Sub", "",
        "Public Sub ClearSink()", "    Set sink = Nothing", "End Sub", "",
        "Public Sub Fire(ByVal n As Long)", "    Dim i As Long", "    For i = 1 To n", "        sink.OnEvent i", "    Next",
        "End Sub", "",
        "Public Sub RaiseError()", '    Err.Raise vbObjectError + 77, "WzProbeLib.Platform", "error raised in the DLL"',
        "End Sub"])
    classes = ["Class=ISink; ISink.cls", "Class=Platform; Platform.cls"]
    for c in range(API_CLASSES):
        body = []
        for k in range(API_PER_CLASS):
            n = c * API_PER_CLASS + k
            body += ["Public Function zq%d(ByVal a As Long) As Long" % n, "    zq%d = a + %d" % (n, n), "End Function"]
        write(folder, "Api%d.cls" % c, cls_header("Api%d" % c, "global") + body)
        classes.append("Class=Api%d; Api%d.cls" % (c, c))
    vbp = ["Type=OleDll", STDOLE] + classes + ["Module=pmod; pmod.bas", 'Startup="(None)"',
                                                'ExeName32="WzProbeLib.dll"', 'Name="WzProbeLib"',
                                                'CompatibleMode="0"', "ThreadingModel=1"]
    with open(os.path.join(folder, "WzProbeLib.vbp"), "w", newline="\r\n") as f:
        f.write("\n".join(vbp) + "\n" + OPTIONS)


HOST_MAIN = r'''Attribute VB_Name = "hostmain"
Option Explicit

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (arr() As Any) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (f As Currency) As Long

Private Type TILE
    texture As Long
    height As Long
    info As Byte
    tileVisBits As Byte
    pad As Integer
    limitedContinent As Long
End Type

Public sinkCount As Long
Private logf As Integer
Private view() As TILE
Private sa As SAFEARRAY1D

Private Sub LogLine(ByVal s As String)
    Print #logf, s
End Sub

Private Function Now_() As Double
    Dim c As Currency, f As Currency
    QueryPerformanceCounter c
    QueryPerformanceFrequency f
    Now_ = CDbl(c) / CDbl(f)
End Function

Sub Main()
    Dim i As Long, s As Long, p As Long, t0 As Double, t1 As Double, probe As TILE
    Dim sk As CSink
    On Error GoTo fail
    logf = FreeFile
    Open App.Path & "\host.log" For Output As #logf

    ' 1. GlobalMultiUse calls (the first one creates the DLL's global object)
    s = PlatAdd(1, 2)
    t0 = Now_()
    For i = 1 To 1000000
        s = PlatAdd(s, 1)
    Next
    t1 = Now_()
    LogLine "1. PlatAdd from the exe: result " & s & " (expect 1000003), " & Format$((t1 - t0) * 1000000# / 1000000#, "0.000") & " us per call"

    ' 2. the DLL's TILE array seen from the exe
    p = MapCreate(4096)
    sa.cDims = 1
    sa.fFeatures = &H11
    sa.cbElements = LenB(probe)
    sa.pvData = p
    sa.cElements = 4096
    CopyMemory ByVal VarPtrArray(view), VarPtr(sa), 4
    LogLine "2. TILE size: DLL " & TileSize() & ", exe " & LenB(probe) & "; view(100) = " & view(100).texture & "/" & view(100).height & "/" & view(100).info & " (expect 100/200/100)"
    MapTouch
    LogLine "   after MapTouch in the DLL: view(100).height = " & view(100).height & " (expect 201), view(4095).height = " & view(4095).height & " (expect 8191)"
    view(7).texture = 777
    LogLine "   written by the exe, read back: " & view(7).texture
    p = 0
    CopyMemory ByVal VarPtrArray(view), p, 4

    ' 3. callbacks from the DLL into the exe
    Set sk = New CSink
    SetSink sk
    t0 = Now_()
    Fire 1000000
    t1 = Now_()
    ClearSink
    LogLine "3. callbacks: " & sinkCount & " (expect 1000000), " & Format$((t1 - t0) * 1000000# / 1000000#, "0.000") & " us per callback"

    ' 4. an error raised in the DLL
    On Error Resume Next
    RaiseError
    LogLine "4. error from the DLL: number vbObjectError+" & (Err.Number - vbObjectError) & ", source " & Err.Source & ", " & Err.Description
    On Error GoTo fail
    LogLine "done"
    Close #logf
    Exit Sub
fail:
    LogLine "VB error " & Err.Number & ": " & Err.Description
    Close #logf
End Sub
'''


def make_host(folder, tlb_ref):
    with open(os.path.join(folder, "hostmain.bas"), "w", newline="") as f:
        f.write(HOST_MAIN.replace("\r\n", "\n").replace("\n", "\r\n"))
    write(folder, "CSink.cls", cls_header("CSink", "private") +
          ["Implements ISink", "", "Private Sub ISink_OnEvent(ByVal i As Long)", "    sinkCount = sinkCount + 1",
           "End Sub"])
    vbp = ["Type=Exe", STDOLE, tlb_ref, "Module=hostmain; hostmain.bas", "Class=CSink; CSink.cls",
           'Startup="Sub Main"', 'ExeName32="host.exe"', 'Name="ProbeHost"']
    with open(os.path.join(folder, "host.vbp"), "w", newline="\r\n") as f:
        f.write("\n".join(vbp) + "\n" + OPTIONS)


def make_names_exe(folder, tlb_ref, own_names, call_api):
    """An exe with about own_names names of its own, calling all DLL api functions or none."""
    counter = [0]
    modules = 100
    vbp = ["Type=Exe", STDOLE, tlb_ref]
    calls = []
    for m in range(modules):
        text, first = names_gen.module_source(m, own_names // modules, counter)
        with open(os.path.join(folder, "m%d.bas" % m), "w", newline="") as f:
            f.write(text)
        vbp.append("Module=m%d; m%d.bas" % (m, m))
        calls.append("    s = s + %s(%d)" % (first, m))
    api = []
    if call_api:
        n = API_CLASSES * API_PER_CLASS
        for c in range(0, n, 500):       # several procedures: 64K of code per procedure
            api += ["Private Function api%d() As Long" % c, "    Dim s As Long"]
            api += ["    s = s + zq%d(1)" % k for k in range(c, min(n, c + 500))]
            api += ["    api%d = s" % c, "End Function", ""]
            calls.append("    s = s + api%d()" % c)
    main = ['Attribute VB_Name = "probemain"', "Option Explicit", ""] + api + [
        "Sub Main()", "    Dim s As Long, f As Integer"] + calls + [
        "    f = FreeFile", '    Open App.Path & "\\names.log" For Output As #f',
        '    Print #f, "ran, checksum " & s', "    Close #f", "End Sub"]
    write(folder, "probemain.bas", main)
    vbp += ["Module=probemain; probemain.bas", 'Startup="Sub Main"', 'ExeName32="names.exe"', 'Name="NamesHost"']
    with open(os.path.join(folder, "names.vbp"), "w", newline="\r\n") as f:
        f.write("\n".join(vbp) + "\n" + OPTIONS)
    return counter[0]


def make(folder, vbp):
    out = os.path.join(folder, "build.txt")
    if os.path.exists(out):
        os.remove(out)
    subprocess.run([VB6, "/make", vbp, "/out", "build.txt"], cwd=folder, timeout=1800)
    with open(out, encoding="latin-1") as f:
        return " ".join(f.read().split())


def reg(key):
    r = subprocess.run(["reg", "query", key, "/reg:32"], capture_output=True, text=True)
    return r.stdout if r.returncode == 0 else ""


def guid(text):
    """The GUID in a reg query's REG_SZ value (the key path above it may hold another one)."""
    m = re.search(r"REG_SZ\s+(\{[0-9A-Fa-f-]{36}\})", text)
    return m.group(1) if m else None


def manifest(folder, dll_path, tlb, clsids):
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">',
             '  <assemblyIdentity type="win32" name="ProbeHost" version="1.0.0.0"/>',
             '  <file name="%s">' % os.path.basename(dll_path)]
    for progid, clsid in clsids.items():
        lines.append('    <comClass clsid="%s" progid="%s" threadingModel="Apartment" tlbid="%s"/>' % (clsid, progid, tlb))
    lines += ['    <typelib tlbid="%s" version="1.0" helpdir="" flags=""/>' % tlb, "  </file>", "</assembly>"]
    with open(os.path.join(folder, "host.exe.manifest"), "w", newline="\r\n") as f:
        f.write("\n".join(lines) + "\n")


def main(work):
    work = os.path.abspath(work)
    old = os.path.join(work, "dll", "WzProbeLib.dll")
    if os.path.exists(old):
        subprocess.run(["regsvr32", "/u", "/s", old])     # a rebuild gets new GUIDs: drop the old ones
    if os.path.isdir(work):
        shutil.rmtree(work)
    dll = os.path.join(work, "dll")
    os.makedirs(dll)
    make_dll(dll)
    result = make(dll, "WzProbeLib.vbp")
    print("DLL:", result)
    if "succeeded" not in result:
        return 1
    clsids = {}
    for c in ["Platform"] + ["Api%d" % i for i in range(API_CLASSES)]:
        clsids["WzProbeLib." + c] = guid(reg("HKCR\\WzProbeLib.%s\\Clsid" % c))
    tlb = guid(reg("HKCR\\CLSID\\%s\\TypeLib" % clsids["WzProbeLib.Platform"])) if clsids["WzProbeLib.Platform"] else None
    versions = re.findall(r"\\(\d+\.\d+)\s*$", reg("HKCR\\TypeLib\\%s" % tlb), re.M) if tlb else []
    print("registered: Platform CLSID %s, type library %s, versions %s" % (clsids["WzProbeLib.Platform"], tlb, versions))
    if not tlb or not versions:
        print("the DLL is not registered: stop")
        return 1
    dll_path = os.path.join(dll, "WzProbeLib.dll")
    tlb_ref = r"Reference=*\G%s#%s#0#%s#WzProbeLib" % (tlb, versions[0], dll_path)

    host = os.path.join(work, "host")
    os.makedirs(host)
    make_host(host, tlb_ref)
    result = make(host, "host.vbp")
    print("host:", result)
    if "succeeded" in result:
        shutil.copy(dll_path, host)          # beside the exe, for the reg-free run
        subprocess.run([os.path.join(host, "host.exe")], cwd=host, timeout=300)
        with open(os.path.join(host, "host.log"), encoding="latin-1") as f:
            print(f.read())

    for own, call in ((29000, True), (29000, False), (27000, True)):
        folder = os.path.join(work, "names_%d_%s" % (own, "calls" if call else "nocalls"))
        os.makedirs(folder)
        n = make_names_exe(folder, tlb_ref, own, call)
        result = make(folder, "names.vbp")
        ran = ""
        if "succeeded" in result:
            subprocess.run([os.path.join(folder, "names.exe")], cwd=folder, timeout=120)
            with open(os.path.join(folder, "names.log"), encoding="latin-1") as f:
                ran = f.read().strip()
        print("5. exe with %d own names, %s the DLL's %d functions: %s %s"
              % (n, "calling" if call else "not calling", API_CLASSES * API_PER_CLASS, result or "(no output)", ran))

    # 6. reg-free: unregister, run host.exe with a manifest naming the DLL beside it
    manifest(host, dll_path, tlb, clsids)
    r = subprocess.run(["regsvr32", "/u", "/s", dll_path])
    print("6. regsvr32 /u exit code %d; Platform still registered: %s" % (r.returncode, bool(reg("HKCR\\WzProbeLib.Platform\\Clsid"))))
    if os.path.exists(os.path.join(host, "host.exe")):
        os.remove(os.path.join(host, "host.log"))
        subprocess.run([os.path.join(host, "host.exe")], cwd=host, timeout=300)
        with open(os.path.join(host, "host.log"), encoding="latin-1") as f:
            print("   host.exe, unregistered DLL, with manifest:\n" + f.read())
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
