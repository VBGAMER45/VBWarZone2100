"""
check_names.py - finds names that clash in VB6 (it ignores case) before the
compiler does, one at a time. QuakeVB's tool, extended to WarzoneVB's
projects (tools/build.py's PROJECTS, each with src/shared).

  python -I tools/check_names.py

Reports, case-insensitively, per project:
* a module-level name (Sub, Function, variable, Const, Type, Enum, Declare)
  defined twice in one module, Private or Public
* a Public name defined in two modules
* a module name equal to a Public name
* a Public name of another module that a module also defines Private
  (legal in VB, but a likely mistake when porting)
* a Public name that is also a global of a DLL the project uses (the
  methods of its GlobalMultiUse API class, its public enums and their
  members): legal, the project's own name wins, but it is the wrong one
  for every caller that meant the DLL's
* a parameter or local variable spelled like a Public name (own or a used
  DLL's) or a name of its own module in another case: C's yaw and YAW are
  two names, VB's one, and inside the procedure the local wins
Enum members count as Public constants. Class modules contribute their
public enums (to the project and, in a DLL, to its users); their methods
are not globals, except a GlobalMultiUse API class's for the DLL's users.
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402  (PROJECTS, SRC, SHARED)

ROOT = build.ROOT

DECL = re.compile(
    r"^\s*(?:(Public|Private|Global|Dim|Friend)\s+)?"
    r"(?:(Const|Declare\s+(?:Sub|Function)|Sub|Function|Property\s+(?:Get|Let|Set)|Type|Enum)\s+)?"
    r"([A-Za-z_]\w*)", re.I)
END_PROC = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
START_PROC = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)", re.I)
TYPE_START = re.compile(r"^\s*(?:(?:Public|Private)\s+)?(Type|Enum)\s+([A-Za-z_]\w*)", re.I)
TYPE_END = re.compile(r"^\s*End\s+(Type|Enum)\b", re.I)
KEYWORDS = {"option", "attribute", "end", "rem", "#if", "#end", "implements", "version", "begin"}
# docs/CONVENTIONS.md, "Names": VB statements, functions and objects that make trouble as names
RESERVED = {"scale", "line", "name", "type", "next", "step", "len", "left", "right", "time", "date", "point",
            "dir", "str", "print", "open", "close", "get", "put", "me", "empty", "string", "debug",
            "erase"}
PARAM_WORDS = {"byval", "byref", "optional", "paramarray"}


def split_decl_list(rest):
    """names in 'a As Long, b() As Byte, c As X' (after the first)"""
    names, depth, part = [], 0, ""
    for ch in rest:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            names.append(part)
            part = ""
        else:
            part += ch
    names.append(part)
    out = []
    for p in names[1:]:
        m = re.match(r"\s*([A-Za-z_]\w*)", p)
        if m:
            out.append(m.group(1))
    return out


def param_names(code, start):
    i = code.find("(", start)
    if i < 0:
        return []
    depth, j = 0, i
    for j in range(i, len(code)):
        if code[j] == "(":
            depth += 1
        elif code[j] == ")":
            depth -= 1
            if depth == 0:
                break
    out = []
    for part in code[i + 1:j].split(","):
        words = [w for w in re.findall(r"[A-Za-z_]\w*", re.split(r"\s+As\s+", part)[0]) if w.lower() not in PARAM_WORDS]
        if words:
            out.append(words[0])
    return out


def scan(path):
    """module name, whether it is a class (and a GlobalMultiUse one), its
    module-level names [(name, public, kind, line)], the parameters and
    locals of its procedures [(name, procedure, line)]"""
    is_class = path.lower().endswith(".cls")
    modname = os.path.splitext(os.path.basename(path))[0]
    names, locals_ = [], []
    proc, in_proc, block, block_public = None, False, None, False
    with open(path, encoding="latin-1") as f:
        lines = f.read().split("\n")
    global_ns = any(l.strip() == "Attribute VB_GlobalNameSpace = True" for l in lines)
    joined, buf = [], ""
    for ln in lines:
        ln = ln.rstrip("\r")
        if buf:
            ln = buf + " " + ln.strip()
            buf = ""
        if ln.rstrip().endswith(" _"):
            buf = ln.rstrip()[:-2]
            continue
        joined.append(ln)
    for no, ln in enumerate(joined, 1):
        code = re.sub(r"\"[^\"]*\"", '""', ln).split("'")[0]
        if code.startswith("Attribute VB_Name"):
            modname = ln.split('"')[1]
            continue
        if block:
            if TYPE_END.match(code):
                block = None
            elif block == "enum":
                m = re.match(r"^\s*([A-Za-z_]\w*)", code)
                if m:
                    names.append((m.group(1), block_public, "enum member", no))
            continue
        if in_proc:
            if END_PROC.match(code):
                in_proc = False
                continue
            m = re.match(r"^\s*(Dim|Static)\s+([A-Za-z_]\w*)", code, re.I)
            if m:
                locals_.append((m.group(2), proc, no))
                for extra in split_decl_list(code[m.end():]):
                    locals_.append((extra, proc, no))
            continue
        m = TYPE_START.match(code)
        if m:
            public = not code.strip().lower().startswith("private")
            block = m.group(1).lower()
            block_public = public
            names.append((m.group(2), public, block, no))
            continue
        m = START_PROC.match(code)
        if m:
            public = not re.match(r"^\s*(Private|Friend)\b", code, re.I)
            # a class's methods are members, not globals; a module's Property Get/Let/Set share their name
            if is_class:
                kind = "method"
            elif m.group(1).lower().startswith("property"):
                kind = "property"
            else:
                kind = "proc"
            names.append((m.group(2), public and not is_class, kind, no))
            in_proc, proc = True, m.group(2)
            for pn in param_names(code, m.end()):
                locals_.append((pn, proc, no))
            continue
        m = DECL.match(code)
        if not m or not m.group(1):
            continue
        scope, kind, name = m.group(1).lower(), (m.group(2) or "var").split()[0].lower(), m.group(3)
        if name.lower() in KEYWORDS:
            continue
        public = scope in ("public", "global") and not is_class
        names.append((name, public, kind, no))
        if kind in ("var", "const"):
            for extra in split_decl_list(code[m.end():]):
                names.append((extra, public, kind, no))
    return modname, is_class, global_ns, names, locals_


def project_files(sub):
    files = []
    dirs = [os.path.join(build.SRC, sub), build.SHARED]
    if sub in build.SHARED_GFX_USERS:
        dirs.append(build.SHARED_GFX)
    for d in dirs:
        if os.path.isdir(d):
            files += [os.path.join(d, n) for n in sorted(os.listdir(d)) if n.lower().endswith((".bas", ".cls", ".frm"))]
    return files


def exported(sub):
    """The globals a DLL gives its users: its GlobalMultiUse classes' public methods, public enums and members."""
    out = {}
    for path in project_files(sub):
        if not path.lower().endswith(".cls"):
            continue
        modname, is_class, global_ns, names, _ = scan(path)
        for name, public, kind, no in names:
            if kind in ("enum", "enum member") and public:
                out[name.lower()] = (os.path.relpath(path, ROOT), no, name)
        if global_ns:
            for name, public, kind, no in names:
                if kind == "method":
                    out.setdefault(name.lower(), (os.path.relpath(path, ROOT), no, name))
    return out


def check_project(name, sub, refs):
    problems = 0
    publics, modules, scanned = {}, {}, []
    imported = {}
    folders = {p[0]: p[2] for p in build.PROJECTS}
    for r in refs:
        for key, v in exported(folders[r]).items():
            imported.setdefault(key, v + (r,))
    for path in project_files(sub):
        modname, is_class, global_ns, names, locals_ = scan(path)
        rel = os.path.relpath(path, ROOT)
        modules[modname.lower()] = rel
        scanned.append((rel, modname, is_class, names, locals_))
        seen = {}
        for nm, public, kind, no in names:
            key = nm.lower()
            if key in seen and kind != "method" and seen[key][2] != "method" and not (kind == "property" and seen[key][2] == "property"):
                print("%s: %s:%d: %s clashes with %s (line %d) in the same module" % (name, rel, no, nm, seen[key][0], seen[key][1]))
                problems += 1
            elif key not in seen:
                seen[key] = (nm, no, kind)
            if public or kind in ("enum", "enum member") and public:
                if key in publics and publics[key][0] != rel:
                    print("%s: %s:%d: Public %s is also Public in %s:%d" % (name, rel, no, nm, publics[key][0], publics[key][1]))
                    problems += 1
                else:
                    publics[key] = (rel, no, nm)
    for key, (rel, no, nm) in publics.items():
        if key in modules:
            print("%s: %s:%d: Public %s has the name of module %s" % (name, rel, no, nm, modules[key]))
            problems += 1
        if key in imported and not rel.lower().endswith(".cls"):
            i = imported[key]
            print("%s: %s:%d: Public %s is also %s's %s (%s:%d)" % (name, rel, no, nm, i[3], i[2], i[0], i[1]))
            problems += 1
    for rel, modname, is_class, names, locals_ in scanned:
        for nm, public, kind, no in names:
            if not public and kind not in ("method", "enum member") and nm.lower() in publics and publics[nm.lower()][0] != rel:
                p = publics[nm.lower()]
                print("%s: %s:%d: Private %s hides Public %s of %s:%d" % (name, rel, no, nm, p[2], p[0], p[1]))
                problems += 1
        own = {}
        for nm, public, kind, no in names:
            own[nm.lower()] = (rel, no, nm)
            if nm.lower() in RESERVED and kind != "enum member":
                print("%s: %s:%d: %s is a reserved name (docs/CONVENTIONS.md)" % (name, rel, no, nm))
                problems += 1
        for nm, proc, no in locals_:
            if nm.lower() in RESERVED:
                print("%s: %s:%d: %s in %s is a reserved name (docs/CONVENTIONS.md)" % (name, rel, no, nm, proc))
                problems += 1
            key = nm.lower()
            other = own.get(key) or publics.get(key) or imported.get(key)
            if other and other[2] != nm:
                print("%s: %s:%d: %s in %s is spelled %s in %s:%d" % (name, rel, no, nm, proc, other[2], other[0], other[1]))
                problems += 1
    return problems


def main():
    problems = 0
    for name, kind, sub, refs in build.PROJECTS:
        if os.path.isdir(os.path.join(build.SRC, sub)):
            problems += check_project(name, sub, refs)
    print("%d problem(s)" % problems)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
