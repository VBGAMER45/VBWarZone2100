"""
gen_widgets.py - writes src/game/widget_classes_vb.bas, the generated half of the widget object model
(docs/CONVENTIONS.md, "Widgets"): the kinds, the class depths, the data accessors, allocation, kind
tests, the virtual dispatch and the WzFunc dispatch. Never edit the output; run this after adding or
changing a widget class, a virtual method or a WzFunc id.

  python -I tools\\gen_widgets.py [source folder]      (default src/game)

What it reads from the folder's .bas and .cls files:

* `' (VB) widget class NAME : PARENT` (`' (VB) widget class WIDGET` for the root): a C++ class of the
  WIDGET hierarchy. Every widget is one VB WIDGET object; its kind says which class it is.
* `' (VB) virtual NAME.method`: a virtual method that class NAME declares (introduces).
* `Public Sub|Function NAME_method(ByVal self As WIDGET[, ...])`: class NAME's implementation of a
  method (its override, when an ancestor declares the method virtual).
* A .cls whose header has `' (VB) data of widget class NAME` (named NAME.cls; a shorter name only when
  NAME has over 31 characters, VB's limit for a class): the class of NAME's own fields; every
  widget whose class is NAME or derives from it gets one (w.ext(DEPTH_NAME)), read with NAME_(w).
* `Public Const FN_xxx As Long = n` in a module that has `Public Function <module>_fn(ByVal f As
  WzFunc, args() As Variant) As Variant`: a WzFunc id (a lambda, a std::function, a function pointer)
  that the module's _fn runs.

What it writes:

* `Enum WIDGET_KIND` (K_NAME), `Const DEPTH_NAME` (WIDGET = 0, each derived class one deeper);
* `NAME_(w) As NAME`: the data object of class NAME of widget w;
* `WIDGET_alloc(kind)`: a new WIDGET of that kind, with a new data object for every class of its chain
  that has one (the constructors, NAME_construct, are the hand-written C++ constructors);
* `WIDGET_isKindOf(w, kind)` (dynamic_cast) and `WIDGET_kindName(kind)`;
* `vt_ROOT_method(self, ...)`: the virtual call of method, declared by class ROOT: a Select Case on
  self.kind calling the most derived implementation in each kind's chain;
* `WIDGET_destructAll(self)`: every NAME_destruct of the chain, the most derived first (C++'s virtual
  destructors);
* `WzFunc_dispatch(f, args)`: the WzFunc ids to their module's _fn.
"""
import os
import re
import sys

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "src", "game")
OUT = os.path.join(src, "widget_classes_vb.bas")

CLASS = re.compile(r"^'\s*\(VB\) widget class (\w+)(?:\s*:\s*(\w+))?\s*$")
VIRTUAL = re.compile(r"^'\s*\(VB\) virtual (\w+)\.(\w+)\s*$")
DATA = re.compile(r"^'\s*\(VB\) data of widget class (\w+)\s*$")
PROC = re.compile(r"^(?:Public|Friend)\s+(Sub|Function)\s+(\w+)\s*\((.*)\)\s*(As\s+[\w.]+)?\s*(?:'.*)?$", re.I)
FN_CONST = re.compile(r"^Public Const (FN_\w+) As Long = (\d+)\b")
# return types a dispatcher assigns with Set
OBJECT_TYPES = ("widget", "object", "wzfunc", "w_screen", "wztext", "collection", "widgethelp", "iv_image")
FN_PROC = re.compile(r"^Public Function (\w+)_fn\s*\(\s*ByVal f As WzFunc\s*,\s*args\(\) As Variant\s*\)\s*As Variant", re.I)

classes = {}        # name -> parent (None for the root)
class_file = {}
virtuals = {}       # (declaring class, method) -> True
procs = {}          # proc name (lower) -> (kind Sub/Function, name, params, return)
data_classes = set()
data_file = {}      # class -> its data class's name (the class's own, but VB's 31-character limit)
fn_ids = {}         # id value -> (const name, module)
fn_modules = {}     # module name -> True when it has <module>_fn
problems = []

for n in sorted(os.listdir(src)):
    if not n.lower().endswith((".bas", ".cls")) or n.lower() == "widget_classes_vb.bas":
        continue
    path = os.path.join(src, n)
    module = None
    lines = open(path, encoding="latin-1").read().replace("\r\n", "\n").split("\n")
    consts = []
    for l in lines:
        if l.startswith("Attribute VB_Name"):
            module = l.split('"')[1]
        m = CLASS.match(l)
        if m:
            if m.group(1) in classes:
                problems.append("class %s declared twice (%s, %s)" % (m.group(1), class_file[m.group(1)], n))
            classes[m.group(1)] = m.group(2)
            class_file[m.group(1)] = n
            continue
        m = VIRTUAL.match(l)
        if m:
            virtuals[(m.group(1), m.group(2))] = True
            continue
        m = DATA.match(l)
        if m and n.lower().endswith(".cls"):
            data_classes.add(m.group(1))
            data_file[m.group(1)] = os.path.splitext(n)[0]
            if os.path.splitext(n)[0] != m.group(1) and len(m.group(1)) <= 31:
                problems.append("%s: the data class of %s must be named %s.cls" % (n, m.group(1), m.group(1)))
            continue
        m = PROC.match(l)
        if m:
            procs[m.group(2).lower()] = (m.group(1), m.group(2), m.group(3), m.group(4) or "")
        m = FN_CONST.match(l)
        if m:
            consts.append((m.group(1), int(m.group(2))))
        m = FN_PROC.match(l)
        if m:
            fn_modules[m.group(1)] = True
    if module and module in fn_modules:
        for name, value in consts:
            if value in fn_ids:
                problems.append("WzFunc id %d is both %s (%s) and %s (%s)" % (value, fn_ids[value][0], fn_ids[value][1], name, module))
            fn_ids[value] = (name, module)

if "WIDGET" not in classes:
    problems.append("no '(VB) widget class WIDGET' declaration")


def chain(name):
    out = []
    while name is not None:
        if name not in classes:
            problems.append("class %s: parent %s is not declared" % (out[-1] if out else "?", name))
            break
        out.append(name)
        name = classes[name]
        if len(out) > 32:
            problems.append("class chain loops at %s" % name)
            break
    return out          # most derived first, WIDGET last


depth = {c: len(chain(c)) - 1 for c in classes}
names = sorted(classes, key=lambda c: (depth[c], c))
out = ['Attribute VB_Name = "widget_classes_vb"',
       "' (VB) generated by tools/gen_widgets.py from the widget class declarations of src/game: do not",
       "' edit. The kinds, depths, data accessors, allocation, kind tests, virtual dispatch and WzFunc",
       "' dispatch of the widget object model (docs/CONVENTIONS.md, \"Widgets\").",
       "Option Explicit", "",
       "Public Enum WIDGET_KIND"]
for i, c in enumerate(names):
    out.append("    K_%s = %d" % (c, i))
out += ["End Enum", ""]
for c in names:
    out.append("Public Const DEPTH_%s As Long = %d" % (c, depth[c]))
out += ["Public Const WIDGET_MAX_DEPTH As Long = %d" % (max(depth.values()) if depth else 0), ""]

# data accessors
for c in names:
    if c in data_classes:
        out += ["Public Function %s_(ByVal w As WIDGET) As %s" % (c, data_file[c]),
                "    Set %s_ = w.ext(DEPTH_%s)" % (c, c),
                "End Function", ""]

# allocation
out += ["' a new WIDGET of kind kind, with the data objects of its chain (the constructors are NAME_construct)",
        "Public Function WIDGET_alloc(ByVal kind As WIDGET_KIND) As WIDGET",
        "    Dim w As WIDGET",
        "    Set w = New WIDGET",
        "    w.kind = kind",
        "    Select Case kind"]
for c in names:
    with_data = [x for x in chain(c) if x in data_classes]
    if with_data:
        out.append("    Case K_%s" % c)
        for x in with_data:
            out.append("        Set w.ext(DEPTH_%s) = New %s" % (x, data_file[x]))
out += ["    End Select", "    Set WIDGET_alloc = w", "End Function", ""]

# kind tests
out += ["' dynamic_cast<KIND *>(w) != nullptr: w's class is kind or derives from it",
        "Public Function WIDGET_isKindOf(ByVal w As WIDGET, ByVal kind As WIDGET_KIND) As Boolean",
        "    If w Is Nothing Then Exit Function",
        "    Select Case kind"]
for c in names:
    derived = [d for d in names if c in chain(d)]
    out.append("    Case K_%s" % c)
    out.append("        Select Case w.kind")
    for i in range(0, len(derived), 12):
        out.append("        Case %s: WIDGET_isKindOf = True" % ", ".join("K_" + d for d in derived[i:i + 12]))
    out.append("        End Select")
out += ["    End Select", "End Function", "",
        "Public Function WIDGET_kindName(ByVal kind As WIDGET_KIND) As String",
        "    Select Case kind"]
for c in names:
    out.append('    Case K_%s: WIDGET_kindName = "%s"' % (c, c))
out += ["    End Select", "End Function", ""]


def call_args(params):
    """the argument list of a call from a parameter list (names only), without self"""
    out = []
    for p in params.split(",")[1:]:
        p = p.strip()
        if not p:
            continue
        w = [x for x in p.replace("()", "").split() if x.lower() not in ("byval", "byref", "optional", "paramarray")]
        name = w[0]
        out.append(name)
    return ", ".join(out)


# virtual dispatch
for (root, method) in sorted(virtuals):
    impl = procs.get(("%s_%s" % (root, method)).lower())
    if root not in classes:
        problems.append("virtual %s.%s: class %s is not declared" % (root, method, root))
        continue
    if impl is None:
        problems.append("virtual %s.%s: no %s_%s(ByVal self As WIDGET, ...) implementation" % (root, method, root, method))
        continue
    kind, pname, params, ret = impl
    if not re.match(r"\s*ByVal self As WIDGET\b", params, re.I):
        problems.append("%s: the first parameter must be ByVal self As WIDGET" % pname)
        continue
    vname = "vt_%s_%s" % (root, method)
    is_fn = kind.lower() == "function"
    out.append("Public %s %s(%s)%s" % (kind, vname, params, (" " + ret) if ret else ""))
    out.append("    Select Case self.kind")
    groups = {}
    for c in names:
        ch = chain(c)
        if root not in ch:
            continue
        for x in ch:
            if ("%s_%s" % (x, method)).lower() in procs:
                groups.setdefault(x, []).append(c)
                break
            if x == root:
                break
    args = call_args(params)
    for x in sorted(groups):
        p2 = procs[("%s_%s" % (x, method)).lower()]
        if p2[0].lower() != kind.lower():
            problems.append("%s_%s is a %s, %s_%s a %s" % (x, method, p2[0], root, method, kind))
        ks = groups[x]
        call = "%s_%s(self%s)" % (x, method, (", " + args) if args else "")
        for i in range(0, len(ks), 12):
            case = "    Case %s" % ", ".join("K_" + k for k in ks[i:i + 12])
            if is_fn and ret.lower().split()[-1] in OBJECT_TYPES:
                out.append(case + ": Set %s = %s" % (vname, call))
            elif is_fn:
                out.append(case + ": %s = %s" % (vname, call))
            else:
                out.append(case + ": %s_%s self%s" % (x, method, (", " + args) if args else ""))
    out += ["    End Select", "End %s" % kind, ""]

# destructors
out += ["' the C++ virtual destructors: every NAME_destruct of self's chain, the most derived first",
        "Public Sub WIDGET_destructAll(ByVal self As WIDGET)",
        "    Select Case self.kind"]
for c in names:
    ds = [x for x in chain(c) if ("%s_destruct" % x).lower() in procs]
    if ds:
        out.append("    Case K_%s" % c)
        for x in ds:
            out.append("        %s_destruct self" % x)
out += ["    End Select", "End Sub", ""]

# WzFunc dispatch
out += ["' a WzFunc's id to its module's _fn",
        "Public Function WzFunc_dispatch(ByVal f As WzFunc, args() As Variant) As Variant",
        "    Select Case f.id"]
by_module = {}
for value, (name, module) in sorted(fn_ids.items()):
    by_module.setdefault(module, []).append(name)
for module in sorted(by_module):
    ids = by_module[module]
    for i in range(0, len(ids), 10):
        out.append("    Case %s" % ", ".join(ids[i:i + 10]))
        out.append("        WzFunc_dispatch_set WzFunc_dispatch, %s_fn(f, args)" % module)
out += ["    Case Else",
        "        debug_ 0, LOG_ERROR, \"WzFunc_dispatch\", \"(VB) no WzFunc %d\", f.id",
        "    End Select", "End Function", "",
        "' (assigns an object or a value)",
        "Private Sub WzFunc_dispatch_set(dst As Variant, ByVal v As Variant)",
        "    If IsObject(v) Then Set dst = v Else dst = v",
        "End Sub"]

text = "\r\n".join(out) + "\r\n"
text.encode("ascii")
with open(OUT, "w", encoding="latin-1", newline="") as f:
    f.write(text)
print("%s: %d classes (%d with data), %d virtuals, %d WzFunc ids" % (OUT, len(classes), len(data_classes), len(virtuals), len(fn_ids)))
for p in problems:
    print("PROBLEM: " + p)
sys.exit(1 if problems else 0)
