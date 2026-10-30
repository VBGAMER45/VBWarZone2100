"""
check_redim.py - a fixed-size array passed to a procedure that ReDims or Erases that parameter: VB raises
run-time error 10 ("This array is fixed or temporarily locked") only when the call runs.

  python -I tools\\check_redim.py [project folder, default: every one]

Finds each procedure's array parameters (`name() As T`, ByRef) that its body ReDims (without Preserve or
with) or Erases, then each call of that procedure (by name, anywhere in the project) whose argument at that
position is a local or module array declared with bounds (`Dim a(0 To 2) As Long`). A call through a
variable of another name or across a DLL is not followed.
"""
import glob
import os
import re
import sys

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
SRC = os.path.join(ROOT, "src")
folders = sys.argv[1:] or [d for d in os.listdir(SRC) if os.path.isdir(os.path.join(SRC, d))]


def logical_lines(path):
    """(line number, code) with continuations joined and comments and strings blanked"""
    buf, start = "", 0
    for no, raw in enumerate(open(path, encoding="latin-1").read().split("\n"), 1):
        l = raw.rstrip("\r")
        code = re.sub(r'"(?:[^"]|"")*"', '""', l)
        q = code.find("'")
        if q >= 0:
            code = code[:q]
        if not buf:
            start = no
        if code.rstrip().endswith(" _"):
            buf += code.rstrip()[:-2] + " "
            continue
        yield start, buf + code
        buf = ""


PROC = re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?(Sub|Function|Property (?:Get|Let|Set))\s+(\w+)\s*\((.*)$", re.I)


def balanced(text):
    """text up to the parenthesis that closes the one before it"""
    depth = 1
    for i, ch in enumerate(text):
        depth += ch == "("
        depth -= ch == ")"
        if depth == 0:
            return text[:i]
    return text
problems = 0
for folder in folders:
    files = glob.glob(os.path.join(SRC, folder, "*.bas")) + glob.glob(os.path.join(SRC, folder, "*.cls"))
    files += glob.glob(os.path.join(SRC, "shared", "*.bas"))
    resizing = {}          # proc name (lower) -> set of parameter positions it ReDims / Erases
    procs = []             # (path, start, end, name, params, body lines)
    for path in files:
        lines = list(logical_lines(path))
        i = 0
        while i < len(lines):
            m = PROC.match(lines[i][1])
            if not m:
                i += 1
                continue
            j = i + 1
            while j < len(lines) and not re.match(r"^\s*End (Sub|Function|Property)\b", lines[j][1], re.I):
                j += 1
            plist = balanced(m.group(3))
            params = [p.strip() for p in plist.split(",")] if plist.strip() else []
            procs.append((path, lines[i][0], m.group(2), params, lines[i + 1:j]))
            for k, p in enumerate(params):
                pm = re.match(r"^(?:Optional\s+)?(?:ByRef\s+)?(\w+)\(\)", p, re.I)
                if pm and not re.match(r"^(?:Optional\s+)?ByVal\b", p, re.I):
                    name = pm.group(1).lower()
                    for _, code in lines[i + 1:j]:
                        if re.search(r"\b(ReDim(?:\s+Preserve)?|Erase)\s+(?:[^:]*,\s*)?%s\s*(\(|,|$|:)" % re.escape(name), code, re.I):
                            resizing.setdefault(m.group(2).lower(), set()).add(k)
                            break
            i = j + 1
    # module-level fixed arrays per file
    for path, start, name, params, body in procs:
        fixed = set()
        for _, code in body:
            for dm in re.finditer(r"\b(?:Dim|Static)\s+(.*)", code, re.I):
                for part in re.split(r",(?![^(]*\))", dm.group(1)):
                    am = re.match(r"\s*(\w+)\s*\(\s*[^)\s][^)]*\)", part)
                    if am:
                        fixed.add(am.group(1).lower())
        if not fixed:
            continue
        for no, code in body:
            for callee, positions in resizing.items():
                for cm in re.finditer(r"\b%s\b\s*(\(?)(.*)" % re.escape(callee), code, re.I):
                    args = cm.group(2)
                    if cm.group(1):
                        depth, end_ = 1, None
                        for ci, ch in enumerate(args):
                            depth += ch == "("
                            depth -= ch == ")"
                            if depth == 0:
                                end_ = ci
                                break
                        args = args[:end_] if end_ is not None else args
                    parts, depth, cur = [], 0, ""
                    for ch in args:
                        if ch == "," and depth == 0:
                            parts.append(cur)
                            cur = ""
                            continue
                        depth += ch == "("
                        depth -= ch == ")"
                        cur += ch
                    parts.append(cur)
                    for k in positions:
                        if k < len(parts) and parts[k].strip().lower() in fixed:
                            problems += 1
                            print("%s:%d: %s passes the fixed array %s to %s, which resizes it" % (
                                os.path.relpath(path, ROOT), no, name, parts[k].strip(), callee))
print("%d problem(s)" % problems)
