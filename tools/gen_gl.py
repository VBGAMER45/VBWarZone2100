"""
gen_gl.py - generates src/gfx/gl_win.bas, the GL entry points of WzGfx (cluster D).

  python -I tools\\gen_gl.py

Reads 3rdparty/glad (glad.h: the prototypes, the constants, the feature blocks; glad.c: the
loaders gladLoadGLLoader / gladLoadGLES2Loader, the version and extension detection, which
function each feature block loads) and the GL calls of lib/ivis_opengl/gfx_api_gl.cpp, and
writes glad's loader for just those functions and constants as a VB module:

* GL 1.0 / 1.1 functions are `Public Declare ... Lib "opengl32"` (opengl32.dll exports them,
  in a core context too); every later one is a `Public Sub/Function glX(...)` wrapper that
  calls the entry point glad_glX through DispCallFunc (oleaut32), with the arguments in
  module-level Variants prepared once (tools/probe/glprobe.bas measured 0.037 us a call).
  A call to a function that was not loaded raises an error (C++ would call a null pointer).
* glad_glX: the entry points (Long; 0 = not loaded), as glad's variables; the C++ tests
  `glX != nullptr` read them.
* GLAD_GL_<feature>: Booleans, set as glad.c sets them.
* Parameter types: pointers, enums and integers are ByVal Long (an address is a VarPtr);
  GLfloat is Single, GLdouble Double; a `const GLchar *` parameter is a String, passed as a
  NUL-terminated ANSI copy (vbNullString = nullptr).
* glad_call2/4/5(fn, ...): call a function pointer with Long arguments (gfx_api_gl.cpp's
  wz_dyn_* pointers).
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "ref", "warzone2100")
GLAD_H = os.path.join(REF, "3rdparty", "glad", "include", "glad", "glad.h")
GLAD_C = os.path.join(REF, "3rdparty", "glad", "src", "glad.c")
USERS = [os.path.join(REF, "lib", "ivis_opengl", "gfx_api_gl.cpp")]
OUT = os.path.join(ROOT, "src", "gfx", "gl_win.bas")

VB_RESERVED = set("""array type string name len error end get put open close print line left right str date time step
next point dir erase empty scale input lock unlock seek event property me mod and or not xor eqv imp is like to as
in on of set let call sub function dim redim const enum global public private friend preserve optional byval byref
paramarray new nothing true false null variant object single double long integer boolean byte currency decimal
attribute begin select case else elseif if then for each while wend do loop until with goto gosub return resume exit
implements raiseevent withevents debug spc tab stop randomize width height option binary random output shared static
local load unload""".split())

# VB is case-insensitive: a GL constant spelled like one of the port's names (the gl_texture class)
# gets a trailing underscore (docs/CONVENTIONS.md, "Names")
CONST_RENAMES = {"GL_TEXTURE": "GL_TEXTURE_", "GL_BUFFER": "GL_BUFFER_"}


def read(path):
    with open(path, "r", encoding="latin-1") as f:
        return f.read()


def used_names(paths):
    funcs, consts = set(), set()
    for p in paths:
        text = read(p)
        # drop the Emscripten-only parts the port leaves out (they are #if defined(__EMSCRIPTEN__) blocks)
        funcs |= set(re.findall(r"\b(gl[A-Z][A-Za-z0-9]*)\b", text))
        consts |= set(re.findall(r"\b(GL_[A-Za-z0-9_]+)\b", text))
    return funcs, consts


def parse_header(text):
    """prototypes {name: (ret, [(ctype, pname)])}, constants {name: value}, feature names"""
    typedefs = {}
    for m in re.finditer(r"typedef\s+(.+?)\s*\(APIENTRYP\s+PFN(GL[A-Z0-9_]+)PROC\)\s*\((.*?)\);", text):
        ret, pfn, args = m.group(1).strip(), m.group(2), m.group(3).strip()
        typedefs[pfn] = (ret, args)
    protos = {}
    for m in re.finditer(r"GLAPI\s+PFN(GL[A-Z0-9_]+)PROC\s+glad_(gl[A-Za-z0-9_]+);", text):
        pfn, name = m.group(1), m.group(2)
        ret, args = typedefs[pfn]
        params = []
        if args not in ("void", ""):
            for a in args.split(","):
                a = a.strip()
                mm = re.match(r"(.*?)([A-Za-z_][A-Za-z0-9_]*)$", a)
                ctype, pname = mm.group(1).strip(), mm.group(2)
                params.append((ctype, pname))
        protos[name] = (ret, params)
    features = set(re.findall(r"#ifndef (GL_[A-Za-z0-9_]+)\n#define \1 1\n", text))
    consts = {}
    for m in re.finditer(r"^#define (GL_[A-Za-z0-9_]+) (0x[0-9A-Fa-f]+|\d+)\s*$", text, re.M):
        name, val = m.group(1), m.group(2)
        if name in features:
            continue
        v = int(val, 16) if val.startswith("0x") else int(val)
        consts.setdefault(name, v)
    return protos, consts, features


def parse_glad_c(text):
    """the load blocks {block: [functions]} and the loaders' orders and extension lists"""
    blocks = {}
    for m in re.finditer(r"static void load_(GL_[A-Za-z0-9_]+)\(GLADloadproc load\) \{\n\tif\(!GLAD_\1\) return;\n(.*?)\n\}", text, re.S):
        blocks[m.group(1)] = re.findall(r"glad_(gl[A-Za-z0-9_]+) = ", m.group(2))

    def loader(name):
        body = re.search(r"int " + name + r"\(GLADloadproc load\) \{(.*?)\n\}", text, re.S).group(1)
        return re.findall(r"load_(GL_[A-Za-z0-9_]+)\(load\);", body)

    def exts(name):
        body = re.search(r"static int " + name + r"\(void\) \{(.*?)\n\}", text, re.S).group(1)
        return re.findall(r"GLAD_(GL_[A-Za-z0-9_]+) = has_ext\(\"(GL_[A-Za-z0-9_]+)\"\);", body)

    def versions(name):
        body = re.search(r"static void " + name + r"\(void\) \{(.*?)\n\}", text, re.S).group(1)
        return re.findall(r"GLAD_(GL_[A-Za-z0-9_]+) = \(major == (\d+) && minor >= (\d+)\) \|\| major > \d+;", body)

    return (blocks, loader("gladLoadGLLoader"), loader("gladLoadGLES2Loader"),
            exts("find_extensionsGL"), exts("find_extensionsGLES2"),
            versions("find_coreGL"), versions("find_coreGLES2"))


def vb_param_name(n):
    return n + "_" if n.lower() in VB_RESERVED else n


def vb_type(ctype):
    c = ctype.replace(" ", "")
    if c == "constGLchar*":
        return "String"
    if "*" in c:
        return "Long"
    base = c.replace("const", "")
    if base in ("GLfloat", "GLclampf"):
        return "Single"
    if base in ("GLdouble", "GLclampd"):
        return "Double"
    if base in ("GLint64", "GLuint64", "GLint64EXT", "GLuint64EXT"):
        raise SystemExit("64-bit parameter not supported: " + ctype)
    return "Long"


def vt_of(vbt):
    return {"Long": "vbLong", "Single": "vbSingle", "Double": "vbDouble", "String": "vbLong"}[vbt]


def vb_hex(v):
    if v > 0xFFFFFFFF:
        return None
    if v >= 0x80000000:
        v -= 0x100000000
        return "&H%X" % (v & 0xFFFFFFFF)
    return "&H%X&" % v


def main():
    header = read(GLAD_H)
    protos, consts, features = parse_header(header)
    blocks, gl_order, gles_order, gl_exts, gles_exts, gl_vers, gles_vers = parse_glad_c(read(GLAD_C))
    ufuncs, uconsts = used_names(USERS)
    funcs = sorted(f for f in ufuncs if f in protos)
    # the loaders need these themselves
    for f in ("glGetString", "glGetStringi", "glGetIntegerv"):
        if f not in funcs:
            funcs.append(f)
    funcs = sorted(set(funcs))
    used_consts = sorted(c for c in uconsts if c in consts)
    for c in ("GL_VERSION", "GL_EXTENSIONS", "GL_NUM_EXTENSIONS"):
        if c not in used_consts:
            used_consts.append(c)
    used_consts = sorted(set(used_consts))
    block_of = {}
    for b, fs in blocks.items():
        for f in fs:
            block_of.setdefault(f, []).append(b)
    legacy = set(blocks.get("GL_VERSION_1_0", [])) | set(blocks.get("GL_VERSION_1_1", []))
    flags = []
    for b in gl_order + gles_order:
        if b not in flags:
            flags.append(b)
    for f, _a, _b in gl_vers + gles_vers:
        if f not in flags:
            flags.append(f)
    for f, _ in gl_exts + gles_exts:
        if f not in flags:
            flags.append(f)
    for f in ("GL_VERSION_1_0", "GL_VERSION_1_1"):
        if f not in flags:
            flags.insert(0, f)

    L = []
    w = L.append
    w('Attribute VB_Name = "gl_win"')
    w("' (VB) generated by tools/gen_gl.py: do not edit. glad (3rdparty/glad: glad.h, glad.c) for the GL")
    w("' functions and constants lib/ivis_opengl/gfx_api_gl.cpp uses: GL 1.0 / 1.1 functions are")
    w("' opengl32.dll's exports, every later one a wrapper that calls its entry point (glad_glX, loaded")
    w("' by gladLoadGLLoader / gladLoadGLES2Loader through the backend's getGLGetProcAddress) with")
    w("' DispCallFunc. Pointers, enums and integers are ByVal Long (an address is a VarPtr), GLfloat is")
    w("' Single, a const GLchar * is a String (an ANSI copy; vbNullString = nullptr). A call to a function")
    w("' that is not loaded raises an error.")
    w("Option Explicit")
    w("")
    w('Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, prgvt As Integer, prgpvarg As Long, pvargResult As Variant) As Long')
    w('Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)')
    w('Private Declare Function lstrlenA Lib "kernel32" (ByVal p As Long) As Long')
    w("Private Const CC_STDCALL As Long = 4")
    w("")
    w("' glad.h's constants")
    for c in used_consts:
        h = vb_hex(consts[c])
        if h is None:
            continue
        w("Public Const %s As Long = %s%s" % (CONST_RENAMES.get(c, c), h, ("   ' glad.h's %s" % c) if c in CONST_RENAMES else ""))
    w("")
    w("' glad's GLVersion and feature flags (int GLAD_GL_<feature>)")
    w("Public GLVersion_major As Long")
    w("Public GLVersion_minor As Long")
    for f in flags:
        w("Public GLAD_%s As Boolean" % f)
    w("")
    w("' glad's entry points (0 = not loaded)")
    for f in funcs:
        w("Public glad_%s As Long" % f)
    w("")
    w("Private a_(0 To 15) As Variant          ' the arguments of a DispCallFunc call")
    w("Private t_(0 To 15) As Integer          ' their VARTYPEs")
    w("Private p_(0 To 15) As Long             ' their addresses")
    w("Private r_ As Variant                   ' the result")
    w("Private prepared As Boolean")
    w("Private glBackend As IGLBackend         ' the loader (GLADloadproc)")
    w("Private max_loaded_major As Long")
    w("Private max_loaded_minor As Long")
    w("Private exts As String                  ' GL_EXTENSIONS (GL < 3)")
    w("Private exts_i() As String              ' glGetStringi(GL_EXTENSIONS, i) (GL 3+)")
    w("Private num_exts_i As Long")
    w("")
    # declares
    for f in funcs:
        if f not in legacy:
            continue
        ret, params = protos[f]
        ps = ", ".join("ByVal %s As %s" % (vb_param_name(n), vb_type(t)) for t, n in params)
        if ret == "void":
            w('Public Declare Sub %s Lib "opengl32" (%s)' % (f, ps))
        else:
            w('Public Declare Function %s Lib "opengl32" (%s) As Long' % (f, ps))
    w("")
    # helpers
    w("Private Sub prepare()")
    w("    Dim i As Long")
    w("    If prepared Then Exit Sub")
    w("    For i = 0 To 15")
    w("        p_(i) = VarPtr(a_(i))")
    w("    Next")
    w("    prepared = True")
    w("End Sub")
    w("")
    w("Private Sub gl_missing(ByVal fn As String)")
    w('    Err.Raise vbObjectError + 513, "gl_win", "GL function not loaded: " & fn')
    w("End Sub")
    w("")
    w("' the NUL-terminated ANSI copy of s for a const GLchar * argument (its address: VarPtr(b(0)))")
    w("Private Function ansiz(ByVal s As String) As Byte()")
    w("    ansiz = StrConv(s & vbNullChar, vbFromUnicode)")
    w("End Function")
    w("")
    w("' (VB) the C string at p (\"\" for NULL)")
    w("Public Function gl_cstr(ByVal p As Long) As String")
    w("    Dim n As Long, b() As Byte")
    w("    If p = 0 Then Exit Function")
    w("    n = lstrlenA(p)")
    w("    If n = 0 Then Exit Function")
    w("    ReDim b(0 To n - 1)")
    w("    CopyMemory b(0), ByVal p, n")
    w("    gl_cstr = StrConv(b, vbUnicode)")
    w("End Function")
    w("")
    for n in (2, 4, 5):
        args = ", ".join("ByVal a%d As Long" % i for i in range(n))
        w("' (VB) calls the GL function at fn with %d Long arguments (gfx_api_gl.cpp's wz_dyn_* pointers)" % n)
        w("Public Sub glad_call%d(ByVal fn As Long, %s)" % (n, args))
        w("    Dim i As Long")
        w("    If fn = 0 Then gl_missing \"(function pointer)\"")
        w("    prepare")
        for i in range(n):
            w("    a_(%d) = a%d" % (i, i))
        w("    For i = 0 To %d" % (n - 1))
        w("        t_(i) = vbLong")
        w("    Next")
        w("    DispCallFunc 0, fn, CC_STDCALL, vbEmpty, %d, t_(0), p_(0), r_" % n)
        w("End Sub")
        w("")
    # wrappers
    for f in funcs:
        if f in legacy:
            continue
        ret, params = protos[f]
        vparams = [(vb_param_name(n), vb_type(t)) for t, n in params]
        ps = ", ".join("ByVal %s As %s" % (n, t) for n, t in vparams)
        isfunc = ret != "void"
        rtype = vb_type(ret) if isfunc else None
        if rtype == "String":
            rtype = "Long"
        w("Public %s %s(%s)%s" % ("Function" if isfunc else "Sub", f, ps, (" As " + rtype) if isfunc else ""))
        strs = [n for n, t in vparams if t == "String"]
        for n in strs:
            w("    Dim b_%s() As Byte" % n)
        w("    If glad_%s = 0 Then gl_missing \"%s\"" % (f, f))
        w("    prepare")
        for i, (n, t) in enumerate(vparams):
            if t == "String":
                w("    If StrPtr(%s) = 0 Then" % n)
                w("        a_(%d) = 0&" % i)
                w("    Else")
                w("        b_%s = ansiz(%s)" % (n, n))
                w("        a_(%d) = VarPtr(b_%s(0))" % (i, n))
                w("    End If")
            else:
                w("    a_(%d) = %s" % (i, n))
            w("    t_(%d) = %s" % (i, vt_of(t)))
        if isfunc:
            rvt = {"Long": "vbLong", "Single": "vbSingle", "Double": "vbDouble"}[rtype]
            w("    DispCallFunc 0, glad_%s, CC_STDCALL, %s, %d, t_(0), p_(0), r_" % (f, rvt, len(vparams)))
            w("    %s = r_" % f)
            w("End Function")
        else:
            w("    DispCallFunc 0, glad_%s, CC_STDCALL, vbEmpty, %d, t_(0), p_(0), r_" % (f, len(vparams)))
            w("End Sub")
        w("")
    # loader
    w("Private Function load_proc(ByVal name_ As String) As Long")
    w("    load_proc = glBackend.getGLGetProcAddress(name_)")
    w("End Function")
    w("")
    w("Private Sub load_block(ByVal block As String)")
    w("    Select Case block")
    for b in gl_order + gles_order + ["GL_VERSION_1_0", "GL_VERSION_1_1"]:
        pass
    seen = set()
    for b in ["GL_VERSION_1_0", "GL_VERSION_1_1"] + gl_order + gles_order:
        if b in seen:
            continue
        seen.add(b)
        fs = [f for f in blocks.get(b, []) if f in funcs]
        w('    Case "%s"' % b)
        w("        If Not GLAD_%s Then Exit Sub" % b)
        for f in fs:
            w('        glad_%s = load_proc("%s")' % (f, f))
    w("    End Select")
    w("End Sub")
    w("")
    w("Private Function get_exts() As Boolean")
    w("    Dim index As Long")
    w("    If max_loaded_major < 3 Then")
    w("        exts = gl_cstr(glGetString(GL_EXTENSIONS))")
    w("    Else")
    w("        num_exts_i = 0")
    w("        glGetIntegerv GL_NUM_EXTENSIONS, VarPtr(num_exts_i)")
    w("        If num_exts_i > 0 Then")
    w("            ReDim exts_i(0 To num_exts_i - 1)")
    w("        Else")
    w("            Exit Function")
    w("        End If")
    w("        For index = 0 To num_exts_i - 1")
    w("            exts_i(index) = gl_cstr(glGetStringi(GL_EXTENSIONS, index))")
    w("        Next")
    w("    End If")
    w("    get_exts = True")
    w("End Function")
    w("")
    w("Private Sub free_exts()")
    w("    Erase exts_i")
    w("    num_exts_i = 0")
    w("End Sub")
    w("")
    w("Private Function has_ext(ByVal ext As String) As Boolean")
    w("    Dim index As Long")
    w("    If max_loaded_major < 3 Then")
    w('        has_ext = InStr(1, " " & exts & " ", " " & ext & " ", vbBinaryCompare) > 0')
    w("    Else")
    w("        For index = 0 To num_exts_i - 1")
    w("            If exts_i(index) = ext Then")
    w("                has_ext = True")
    w("                Exit Function")
    w("            End If")
    w("        Next")
    w("    End If")
    w("End Function")
    w("")
    for kind, lst in (("GL", gl_exts), ("GLES2", gles_exts)):
        w("Private Function find_extensions%s() As Boolean" % kind)
        w("    If Not get_exts() Then Exit Function")
        for flag, s in lst:
            w('    GLAD_%s = has_ext("%s")' % (flag, s))
        w("    free_exts")
        w("    find_extensions%s = True" % kind)
        w("End Function")
        w("")
    for kind, vers, cap in (("GL", gl_vers, (3, 3)), ("GLES2", gles_vers, (3, 2))):
        w("Private Sub find_core%s()" % kind)
        w("    Dim i As Long, major As Long, minor As Long, version As String, prefixes As Variant, p As Long")
        w('    prefixes = Array("OpenGL ES-CM ", "OpenGL ES-CL ", "OpenGL ES ")')
        w("    version = gl_cstr(glGetString(GL_VERSION))")
        w("    If Len(version) = 0 Then Exit Sub")
        w("    For i = 0 To 2")
        w("        If Left$(version, Len(prefixes(i))) = prefixes(i) Then")
        w("            version = Mid$(version, Len(prefixes(i)) + 1)")
        w("            Exit For")
        w("        End If")
        w("    Next")
        w("    ' sscanf(version, \"%d.%d\", &major, &minor)")
        w("    p = 1")
        w("    Do While p <= Len(version) And Mid$(version, p, 1) Like \"#\"")
        w("        major = major * 10 + Asc(Mid$(version, p, 1)) - 48")
        w("        p = p + 1")
        w("    Loop")
        w("    If Mid$(version, p, 1) = \".\" Then")
        w("        p = p + 1")
        w("        Do While p <= Len(version) And Mid$(version, p, 1) Like \"#\"")
        w("            minor = minor * 10 + Asc(Mid$(version, p, 1)) - 48")
        w("            p = p + 1")
        w("        Loop")
        w("    End If")
        w("    GLVersion_major = major: GLVersion_minor = minor")
        w("    max_loaded_major = major: max_loaded_minor = minor")
        for flag, ma, mi in vers:
            w("    GLAD_%s = (major = %s And minor >= %s) Or major > %s" % (flag, ma, mi, ma))
        w("    If GLVersion_major > %d Or (GLVersion_major >= %d And GLVersion_minor >= %d) Then" % (cap[0], cap[0], cap[1]))
        w("        max_loaded_major = %d" % cap[0])
        w("        max_loaded_minor = %d" % cap[1])
        w("    End If")
        w("End Sub")
        w("")
    for kind, order in (("GL", gl_order), ("GLES2", gles_order)):
        w("' int gladLoad%sLoader(GLADloadproc load): the backend's getGLGetProcAddress is the loader" % kind)
        w("Public Function gladLoad%sLoader(ByVal backend As IGLBackend) As Boolean" % kind)
        w("    Set glBackend = backend")
        w("    prepare")
        w("    GLVersion_major = 0: GLVersion_minor = 0")
        w('    glad_glGetString = load_proc("glGetString")')
        w("    If glad_glGetString = 0 Then Exit Function")
        w("    If glGetString(GL_VERSION) = 0 Then Exit Function")
        w("    find_core%s" % kind)
        exts_started = False
        for b in order:
            if not exts_started and not re.match(r"GL_(ES_)?VERSION_", b):
                w("    If Not find_extensions%s() Then Exit Function" % kind)
                exts_started = True
            w('    load_block "%s"' % b)
        if not exts_started:
            w("    If Not find_extensions%s() Then Exit Function" % kind)
        w("    gladLoad%sLoader = GLVersion_major <> 0 Or GLVersion_minor <> 0" % kind)
        w("End Function")
        w("")
    data = "\r\n".join(L) + "\r\n"
    data.encode("ascii")
    with open(OUT, "w", encoding="ascii", newline="") as f:
        f.write(data)
    print("%s: %d functions (%d from opengl32), %d constants, %d flags" % (OUT, len(funcs), len([f for f in funcs if f in legacy]), len(used_consts), len(flags)))
    missing = sorted(f for f in ufuncs if f not in protos)
    print("names in the sources that glad does not define (not functions):", " ".join(missing))


if __name__ == "__main__":
    main()
