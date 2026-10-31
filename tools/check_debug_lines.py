"""Check the C++ line numbers and function names that the VB log calls carry.

  python -I tools\\check_debug_lines.py [warzone source dir]   (default: ref\\warzone2100)

The VB calls carry what the C++ macros take from __LINE__ and __FUNCTION__
(docs/CONVENTIONS.md, "Logging and asserts"):
    debug_ 509, LOG_WZ, "initialize_ConfigDir", "Write dir: %s", d
    ASSERT_FAILURE "x > 0", "src/droid.cpp:123", "droidUpdate", "bad %d", x
    ASSERT_HELPER_FAILURE 3909, "x > 0", location, function_, "bad %d", x
    printLog_ pCustomLogger, 68, LogLevel_Info, "loadMapData_Internal", "Loading: %s", path
(the last is wzmaplib's debug(logger, level, ...), map_internal_c).
For each call this finds the C++ file (the source paths named in the
module's header comment, before Option Explicit, or the ASSERT's location), and checks that the
line holds a log call with the same format string, inside a function of
that name. Line 0 marks a VB-only log line and is skipped, and so are the
stubs_ modules.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")

VB_STRING = r'"((?:[^"]|"")*)"'
CALL = re.compile(r'\b(debug_|debug__|debug_multiline|wz_info)\s+(\d+)\s*,\s*(?:(\w+)\s*,\s*)?' + VB_STRING +
                  r'(?:\s*,\s*' + VB_STRING + ')?')
MAPCALL = re.compile(r'\b(printLog_)\s+\w+\s*,\s*(\d+)\s*,\s*(\w+)\s*,\s*' + VB_STRING +
                     r'(?:\s*,\s*' + VB_STRING + ')?')
# ASSERT_HELPER with a passed-on location and function (checkDroid): the line of the ASSERT_HELPER call
HELPER = re.compile(r'\bASSERT_HELPER_FAILURE\s+(\d+)\s*,\s*' + VB_STRING + r'\s*,\s*(?:' + VB_STRING + r'|\w+)\s*,\s*\w+\s*,\s*' + VB_STRING)
ASSERT = re.compile(r'\bASSERT_FAILURE\s+' + VB_STRING + r'\s*,\s*"([^":]+):(\d+)"\s*,\s*' + VB_STRING +
                    r'(?:\s*,\s*' + VB_STRING + ')?')
CPP_PATH = re.compile(r'\b((?:lib|src)/[\w/]+\.(?:cpp|h|ypp|lpp))\b')
NAME_PREFIX = re.compile(r'^%(?:name-prefix=|option prefix=|define api\.prefix \{?)"?(\w+?)"?\}?\s*$')
FUNC_DEF = re.compile(r'^[A-Za-z_][\w:<>,\*&\t ]*?[\s\*&]([A-Za-z_]\w*(?:::~?\w+)*)\s*\(')
CTOR_DEF = re.compile(r'^([A-Za-z_]\w*::~?[A-Za-z_]\w*)\s*\(')   # Class::Class( and Class::~Class(
NAME_DEF = re.compile(r'^([A-Za-z_]\w*)\s*\(')                     # a name starting its line, the return type on the line before
TYPE_LINE = re.compile(r'^[A-Za-z_][\w:<>,\*& ]*[\w\*&]\s*$')
MACRO_DEF = re.compile(r'^\s*#\s*define\s+(\w+)\(([^)]*)\)\s*(.*)$')


PRI_MACRO = re.compile(r'%"\s*PRI([a-zA-Z])(?:8|16|32|64|MAX|PTR)\s*"')
# a macro that ends the format: "...%" PRIu32, args
PRI_MACRO_END = re.compile(r'%"\s*PRI([a-zA-Z])(?:8|16|32|64|MAX|PTR)(\s*[,)])')


def vb_to_c(s):
    """A VB string literal's text as it appears inside a C literal."""
    return s.replace('""', '\\"')


def c_formats(text):
    """C source with the <inttypes.h> macros in format strings expanded: %" PRIu32 " is %u."""
    text = PRI_MACRO.sub(lambda m: "%" + m.group(1), text)
    return PRI_MACRO_END.sub(lambda m: "%" + m.group(1) + '"' + m.group(2), text)


def enclosing_function(lines, n):
    # a bison (.ypp) or flex (.lpp) source: the rules' actions run in <prefix>parse / <prefix>lex,
    # and yyerror and the other yy names are renamed with the prefix (%name-prefix, %option prefix)
    prefix = None
    rules = None
    for i, l in enumerate(lines):
        m = NAME_PREFIX.match(l.strip())
        if m and prefix is None:
            prefix = m.group(1)
        if l.strip() == "%%" and rules is None:
            rules = (i + 1, "parse" if any(l2.startswith("%token") for l2 in lines) else "lex")
        elif l.strip() == "%%" and rules is not None and len(rules) == 2:
            rules = (rules[0], rules[1], i + 1)
    if prefix is not None and rules is not None and n > rules[0] and (len(rules) == 2 or n < rules[2]):
        return prefix + rules[1]
    in_comment = block_comment_lines(lines)
    for i in range(n - 1, -1, -1):
        if in_comment[i]:
            continue        # a line inside /* */ (sequence.cpp's "buffer the compressed data (ie, let")
        m = FUNC_DEF.match(lines[i]) or CTOR_DEF.match(lines[i])
        if not m and i > 0 and TYPE_LINE.match(lines[i - 1]) and not lines[i - 1].lstrip().startswith(("#", "//", "return")):
            m = NAME_DEF.match(lines[i])
        if not m and lines[i][:1] not in (" ", "\t", "#", "/") and _body_follows(lines, i):
            # a member defined with a return type FUNC_DEF cannot read (gfx_api_gl.cpp's
            # typename std::pair<..., std::function<void(const void*, size_t)>> Class::name())
            m = QUALIFIED_DEF.match(lines[i])
        if not m and lines[i][:1] in (" ", "\t") and not (i > 0 and lines[i - 1].rstrip().endswith((",", "(", "&&", "||"))):
            # a definition inside a namespace, class or struct (quickjs_backend.cpp's unbox<T>::operator(),
            # box overloads, JSValueToInt32): indented, so only when a body follows its parameter list
            m2 = FUNC_DEF.match(lines[i].lstrip())
            if not m2:
                m2 = INLINE_CTOR_DEF.match(re.sub(r"^(?:virtual|explicit|inline)\s+", "", lines[i].lstrip()))
                if m2 and m2.group(1).lstrip("~") not in class_names(lines):
                    m2 = None
            if m2 and m2.group(1).lstrip("~") not in CPP_STATEMENT_WORDS and not lines[i].lstrip().startswith(CPP_STATEMENT_STARTS) \
                    and _body_follows(lines, i) and not _body_closed_before(lines, i, n - 1):
                # (a local struct's inline constructor whose body closed above line n is not it:
                # textdraw.cpp's glyphRaster in drawText)
                m = m2
        if m and not lines[i].rstrip().endswith(";"):
            name = m.group(1).split("::")[-1]
            if name == "operator":
                name = "operator()"              # GCC's __FUNCTION__ of a call operator
            if prefix is not None and name.startswith("yy"):
                name = prefix + name[2:]
            return name
    return None


_class_cache = {}
_comment_cache = {}


def block_comment_lines(lines):
    """For each line, whether it starts inside a /* */ comment (strings and // comments are skipped)."""
    key = id(lines)
    if key not in _comment_cache:
        flags = []
        inside = False
        for l in lines:
            flags.append(inside)
            k = 0
            quote = None
            while k < len(l):
                if inside:
                    e = l.find("*/", k)
                    if e < 0:
                        break
                    inside = False
                    k = e + 2
                elif quote:
                    if l[k] == "\\":
                        k += 2
                        continue
                    if l[k] == quote:
                        quote = None
                    k += 1
                elif l.startswith("//", k):
                    break
                elif l.startswith("/*", k):
                    inside = True
                    k += 2
                else:
                    if l[k] in "\"'":
                        quote = l[k]
                    k += 1
        _comment_cache[key] = flags
    return _comment_cache[key]


def class_names(lines):
    """The names of the classes and structs a C++ file declares."""
    key = id(lines)
    if key not in _class_cache:
        _class_cache[key] = set(re.findall(r"\b(?:class|struct)\s+([A-Za-z_]\w*)", "\n".join(lines)))
    return _class_cache[key]


INLINE_CTOR_DEF = re.compile(r'^(~?[A-Za-z_]\w*)\s*\(')        # a constructor or destructor in its class's body
QUALIFIED_DEF = re.compile(r'^[A-Za-z_].*?[\s>\*&](\w+::~?\w+)\s*\([^;]*$')
CPP_STATEMENT_WORDS = {"if", "while", "for", "switch", "return", "sizeof", "catch", "defined", "decltype", "alignof"}
CPP_STATEMENT_STARTS = ("return ", "else ", "new ", "delete ", "case ", "throw ", "#", "//", "co_return ")


def _body_follows(lines, i):
    """Whether the parameter list opened on line i (0-based) is followed by a function body ("{" before any ";")."""
    depth, opened = 0, False
    for k in range(i, min(len(lines), i + 12)):
        text = re.sub(r'"(?:\\.|[^"\\])*"', '""', re.sub(r"//.*", "", lines[k]))
        for j, ch in enumerate(text):
            if ch == "(":
                depth, opened = depth + 1, True
            elif ch == ")":
                depth -= 1
                if opened and depth == 0:
                    after = text[j + 1:] + " " + " ".join(re.sub(r"//.*", "", ln) for ln in lines[k + 1:k + 3])
                    a = after.find("{")
                    s = after.find(";")
                    return a >= 0 and (s < 0 or a < s)
    return False


def _body_closed_before(lines, i, n):
    """Whether the body that the definition on line i (0-based) opens closes before line n (0-based)."""
    depth, opened = 0, False
    for k in range(i, n):
        text = re.sub(r"'(?:\\.|[^'\\])'", "''", re.sub(r'"(?:\\.|[^"\\])*"', '""', re.sub(r"//.*", "", lines[k])))
        for ch in text:
            if ch == "{":
                depth, opened = depth + 1, True
            elif ch == "}":
                depth -= 1
                if opened and depth == 0:
                    return True
    return False


_macro_cache = {}

# function-like macros of lib/framework/debug.h that wrap a log call, as file_macros gives them
HEADER_MACROS = {
    "ASSERT_NOT_NULLPTR_OR_RETURN": (["retval", "ptrexpr"],
                                     'ASSERT_OR_RETURN(retval, (ptrexpr) != nullptr, "%s should not be nullptr", #ptrexpr)'),
    # lib/netplay/netplay.h
    "ASSERT_HOST_ONLY": (["failAction"],
                         'if (!NetPlay.isHost) { ASSERT(false, "Host only routine detected for client!"); failAction; }'),
}


def file_macros(lines):
    """The function-like macros of a C++ file: name -> (parameters, body)."""
    key = id(lines)
    if key not in _macro_cache:
        out = {}
        i = 0
        while i < len(lines):
            m = MACRO_DEF.match(lines[i])
            if m:
                body = m.group(3)
                while body.rstrip().endswith("\\") and i + 1 < len(lines):
                    i += 1
                    body = body.rstrip()[:-1] + " " + lines[i].strip()
                out[m.group(1)] = ([p.strip() for p in m.group(2).split(",")], body)
            i += 1
        _macro_cache[key] = out
    return _macro_cache[key]


def _expand_first_macro(text, macros):
    """text with its first use of one of the function-like macros expanded (arguments read with their
    parentheses balanced and string literals skipped), or None when it uses none."""
    best = None
    for name, (params, body) in macros:
        m = re.search(r"\b" + re.escape(name) + r"\s*\(", text)
        if m and (best is None or m.start() < best[0].start()):
            best = (m, params, body)
    if best is None:
        return None
    m, params, body = best
    depth, args, cur, i, end = 1, [], "", m.end(), len(text)
    while i < len(text):
        ch = text[i]
        if ch == '"':
            j = i + 1
            while j < len(text) and text[j] != '"':
                j += 2 if text[j] == "\\" else 1
            cur += text[i:j + 1]
            i = j + 1
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                end = i + 1
                break
        if ch == "," and depth == 1:
            args.append(cur.strip())
            cur = ""
        else:
            cur += ch
        i += 1
    args.append(cur.strip())
    out = body
    if params and params[-1] == "...":
        # a variadic macro (wzapi.cpp's SCRIPT_ASSERT): __VA_ARGS__ is the arguments from the "..." on
        rest = ", ".join(args[len(params) - 1:])
        out = re.sub(r"\b__VA_ARGS__\b", lambda _m: rest, out)
        params = params[:-1]
    for p, a in zip(params, args):
        out = re.sub(r"#\s*" + re.escape(p) + r"\b", lambda _m, a=a: '"' + a + '"', out)
        out = re.sub(r"\b" + re.escape(p) + r"\b", lambda _m, a=a: a, out)
    return text[:m.start()] + out + text[end:]


def macro_expansion(lines, n):
    """The expansion of a function-like macro of the file used on line n
    (ALLOC_STATS, ASSERT_PLAYER_OR_RETURN) whose body holds the log call:
    __LINE__ and __FUNCTION__ are the use's. The call may span lines, and a
    macro may use another (wzapi.cpp's SCRIPT_ASSERT_PLAYER and
    SCRIPT_ASSERT): they are expanded until none is left. None when line n
    uses none."""
    macros = list(file_macros(lines).items()) + list(HEADER_MACROS.items())
    line = lines[n - 1]
    starts = [m.start() for m in (re.search(r"\b" + re.escape(name) + r"\s*\(", line) for name, _ in macros) if m]
    if not starts:
        return None
    # the first macro call on line n, to its closing parenthesis (it may span lines; // comments dropped)
    text = "".join(re.sub(r'("(?:\\.|[^"\\])*")|//.*', lambda m: m.group(1) or "", ln) + "\n"
                   for ln in lines[n - 1:min(len(lines), n + 15)])[min(starts):]
    depth, i = 0, 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            j = i + 1
            while j < len(text) and text[j] != '"':
                j += 2 if text[j] == "\\" else 1
            i = j + 1
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                text = text[:i + 1]
                break
        i += 1
    for _ in range(8):
        expanded = _expand_first_macro(text, macros)
        if expanded is None:
            break
        text = expanded
    return re.sub(r'(?<!\\)"\s*"', "", c_formats(text))       # adjacent literals joined (not an escaped \" before a literal's end)


def call_lines(lines, n):
    """The lines from n (1-based) to the end of the macro call starting there (its parentheses balanced),
    at least three: GCC's __LINE__ is the macro name's line, whatever the lines the call spans."""
    depth, opened, end = 0, False, n
    for i in range(n - 1, min(len(lines), n + 15)):
        for ch in re.sub(r'"(?:\\.|[^"\\])*"', '""', lines[i]):
            if ch == "(":
                depth, opened = depth + 1, True
            elif ch == ")":
                depth -= 1
        end = i + 1
        if opened and depth <= 0:
            break
    # a // comment inside the call is no part of what the macros stringify
    return "".join(re.sub(r'("(?:\\.|[^"\\])*")|//.*', lambda m: m.group(1) or "", ln) + "\n"
                   for ln in lines[n - 1:max(end, n + 2)])


def check_call(lines, n, fn, fmt, wanted):
    """Problems with a call claimed at line n (1-based)."""
    if n < 1 or n > len(lines):
        return ["line %d is beyond the end of the file" % n]
    window = re.sub(r'(?<!\\)"\s+"', "", c_formats(call_lines(lines, n)))     # adjacent literals joined
    problems = []
    expansion = macro_expansion(lines, n)
    if expansion is not None and re.search(wanted, expansion):
        window = expansion
    elif not re.search(wanted, lines[n - 1]):
        problems.append("line %d has no %s: %s" % (n, wanted.replace("\\", ""), lines[n - 1].strip()))
    if problems:
        pass
    elif fmt is not None and vb_to_c(fmt) not in window:
        problems.append("line %d does not have the format \"%s\": %s" % (n, fmt, lines[n - 1].strip()))
    actual = enclosing_function(lines, n)
    if fn and actual and actual != fn:
        problems.append("line %d is in %s, not %s" % (n, actual, fn))
    return problems


OBJ_MACRO = re.compile(r'^\s*#\s*define\s+([A-Za-z_]\w*)\s+([^\\\n]*?)\s*(?://.*|/\*.*)?$')


# <stdint.h>'s limits (MinGW-w64's, as g++ builds statsref): lib/framework/types.h only defines its own
# (UINT8_MAX (255), ...) #ifndef them
STDINT_MACROS = {"INT8_MIN": "(-128)", "INT16_MIN": "(-32768)", "INT32_MIN": "(-2147483647 - 1)",
                 "INT8_MAX": "127", "INT16_MAX": "32767", "INT32_MAX": "2147483647",
                 "UINT8_MAX": "255", "UINT16_MAX": "65535", "UINT32_MAX": "0xffffffffU"}


def object_macros(ref):
    """The object-like macros of the C++ headers: name -> value (the first definition)."""
    out = dict(STDINT_MACROS)
    for sub in ("lib", "src"):
        for folder, _, files in os.walk(os.path.join(ref, sub)):
            for f in files:
                if not f.endswith(".h"):
                    continue
                for line in open(os.path.join(folder, f), encoding="utf-8", errors="replace"):
                    m = OBJ_MACRO.match(line)
                    if m and m.group(2) and '"' not in m.group(2):
                        out.setdefault(m.group(1), m.group(2))
    # glad's GL entry points (#define glX glad_glX), which an ASSERT in gfx_api_gl.cpp stringifies
    glad = os.path.join(ref, "3rdparty", "glad", "include", "glad", "glad.h")
    if os.path.exists(glad):
        for line in open(glad, encoding="utf-8", errors="replace"):
            m = re.match(r"^#define (gl[A-Z]\w*) (glad_gl[A-Z]\w*)\s*$", line)
            if m:
                out.setdefault(m.group(1), m.group(2))
    return out


def file_object_macros(lines):
    """The object-like macros a .cpp file defines itself (track.cpp's MAX_TRACKS): name -> value."""
    out = {}
    for line in lines:
        m = OBJ_MACRO.match(line)
        if m and m.group(2) and '"' not in m.group(2):
            out.setdefault(m.group(1), m.group(2))
    return out


def expand_macros(text, macros):
    """text with the object-like macros replaced, as the preprocessor rescans them."""
    for _ in range(5):
        new = re.sub(r"\b[A-Za-z_]\w*\b", lambda m: macros.get(m.group(0), m.group(0)), text)
        if new == text:
            break
        text = new
    return text


def main(ref):
    cache = {}
    total = bad = 0
    macros = object_macros(ref)

    def cpp(path):
        if path not in cache:
            full = os.path.join(ref, path)
            cache[path] = open(full, encoding="utf-8", errors="replace").read().split("\n") if os.path.exists(full) else None
        return cache[path]

    for folder, _, files in os.walk(SRC):
        for f in sorted(files):
            if not f.lower().endswith((".bas", ".cls")) or f.lower().startswith("stubs_"):
                continue
            path = os.path.join(folder, f)
            text = open(path, encoding="latin-1").read()
            head = text.split("Option Explicit")[0] if "Option Explicit" in text else "\n".join(text.split("\n")[:20])
            header = "\n".join(l for l in head.split("\n") if l.lstrip().startswith("'"))
            sources = CPP_PATH.findall(header)
            for lineno, line in enumerate(text.split("\n"), 1):
                code = line.split("' ")[0] if not line.lstrip().startswith("'") else ""
                for m in list(CALL.finditer(code)) + list(MAPCALL.finditer(code)):
                    kind, n, part, a, b = m.group(1), int(m.group(2)), m.group(3), m.group(4), m.group(5)
                    if n == 0:
                        continue
                    fn, fmt = (a, b) if kind != "wz_info" else (a, b)
                    if kind == "debug_multiline":
                        fmt = None
                    total += 1
                    found = []
                    for s in sources:
                        lines = cpp(s)
                        if lines is None:
                            continue
                        probs = check_call(lines, n, fn, fmt, r"\b(debug|wz_info|debug_multiline|_debug)\s*\(")
                        if not probs:
                            found = []
                            break
                        found.append("%s: %s" % (s, "; ".join(probs)))
                    if not sources:
                        found = ["no C++ source named in the module's header comment"]
                    if found:
                        bad += 1
                        print("%s:%d: %s %d %s" % (os.path.relpath(path, ROOT), lineno, kind, n, fn))
                        for p in found:
                            print("    " + p)
                for m in HELPER.finditer(code):
                    n, expr, fmt = int(m.group(1)), m.group(2).replace('""', '"'), m.group(4)
                    total += 1
                    probs = []
                    for s in sources:
                        lines = cpp(s)
                        if lines is None:
                            continue
                        window = call_lines(lines, n)
                        probs = []
                        if not re.match(r"\s*ASSERT_HELPER\s*\(", lines[n - 1]):
                            probs.append("line %d is no ASSERT_HELPER" % n)
                        elif re.sub(r"\s", "", expr) not in re.sub(r"\s", "", window):
                            probs.append("line %d does not test %s" % (n, expr))
                        elif '"' + vb_to_c(fmt) not in c_formats(window):
                            probs.append("line %d has not the format %s" % (n, fmt))
                        if not probs:
                            break
                    if probs:
                        bad += 1
                        print("%s:%d: ASSERT_HELPER_FAILURE %d" % (os.path.relpath(path, ROOT), lineno, n))
                        for p in probs:
                            print("    " + p)
                for m in ASSERT.finditer(code):
                    expr, s, n, fn, fmt = m.group(1), m.group(2), int(m.group(3)), m.group(4), m.group(5)
                    total += 1
                    lines = cpp(s)
                    probs = ["%s not found" % s] if lines is None else check_call(lines, n, fn, fmt, r"\bASSERT")
                    # the expression is C++ code, not a literal: VB's "" is a plain "
                    c_expr = expr.replace('""', '"')
                    window = call_lines(lines, n) + (macro_expansion(lines, n) or "") if lines is not None else ""
                    # ASSERT passes its expression on to ASSERT_HELPER, which stringifies it: the macros in it are expanded
                    # by then (ASSERT_OR_RETURN stringifies its own parameter, as written)
                    plain_assert = lines is not None and re.search(r"\bASSERT\s*\(", (macro_expansion(lines, n) or lines[n - 1]))
                    if lines is not None and not probs and plain_assert:
                        all_macros = dict(macros)
                        all_macros.update(file_object_macros(lines))   # the .cpp's own object-like macros too
                        if expand_macros(c_expr, all_macros) != c_expr:
                            probs = ["ASSERT prints its expression with the macros expanded: %s" % expand_macros(c_expr, all_macros).replace('"', '""')]
                        elif re.sub(r"\s", "", c_expr) not in re.sub(r"\s", "", expand_macros(window, all_macros)):
                            probs = ["line %d does not test %s" % (n, expr)]
                    elif lines is not None and not probs and re.sub(r"\s", "", c_expr) not in re.sub(r"\s", "", window):
                        probs = ["line %d does not test %s" % (n, expr)]
                    if probs:
                        bad += 1
                        print("%s:%d: ASSERT_FAILURE %s:%d %s" % (os.path.relpath(path, ROOT), lineno, s, n, fn))
                        for p in probs:
                            print("    " + p)
    print("%d log calls checked, %d with problems" % (total, bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(ROOT, "ref", "warzone2100")))
