"""
gen_celt_tables.py - writes src/sound/celt_tables_vb.bas (cluster K, fork O2): the constant tables of
libopus 1.5.2's CELT decoder, taken from the C text (VB6 has no array initialisers):

* celt/static_modes_float.h (the float build's static 48 kHz / 960 mode): window120, logN400,
  cache_index50, cache_bits50, cache_caps50, fft_twiddles48000_960 (as two arrays, .r and .i),
  fft_bitrev480/240/120/60, mdct_twiddles960
* celt/modes.c: eband5ms, band_allocation
* celt/cwrs.c: CELT_PVQ_U_DATA (CUSTOM_MODES undefined; uint32 values kept in a Double array)
* celt/quant_bands.c: e_prob_model (flattened [4][2][42]), eMeans (the float table)

Floats are stored as the bits of C's correctly rounded float literal (hex), decoded at init, so the
VB tables equal the C ones bit for bit. The VB module fills Public arrays in celt_tables_init (called once
by celt_modes_c's opus_custom_mode_create / the decoder's init).

  python -I tools\\gen_celt_tables.py
"""
import os
import re
import struct
import sys
from fractions import Fraction

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CELT = os.path.join(ROOT, "ref", "opus", "celt")
OUT = os.path.join(ROOT, "src", "sound", "celt_tables_vb.bas")


def preprocess(text, defined=()):
    """Drops the lines of #if/#ifdef/#ifndef branches that are not taken (only the conditions these files
    use around the tables: defined(X) / X for the names in `defined`; anything else counts as false except
    the include guards #ifndef X, which count as true)."""
    out = []
    stack = []          # (taking, any_taken)
    for line in text.split("\n"):
        s = line.strip()
        m = re.match(r"#\s*(if|ifdef|ifndef|elif|else|endif)\b(.*)", s)
        if m:
            kw, rest = m.group(1), m.group(2).strip()
            parent = all(t for t, _ in stack)
            if kw in ("if", "ifdef", "ifndef"):
                if kw == "ifdef":
                    cond = rest in defined
                elif kw == "ifndef":
                    cond = rest not in defined
                else:
                    names = re.findall(r"defined\s*\(\s*(\w+)\s*\)", rest)
                    if names and re.fullmatch(r"(!?\s*defined\s*\(\s*\w+\s*\)\s*(&&|\|\|)?\s*)+", rest):
                        # evaluate a simple && / || chain of defined()
                        expr = re.sub(r"defined\s*\(\s*(\w+)\s*\)", lambda mm: str(mm.group(1) in defined), rest)
                        expr = expr.replace("&&", " and ").replace("||", " or ").replace("!", " not ")
                        cond = bool(eval(expr))
                    elif rest == "0":
                        cond = False
                    elif rest == "1":
                        cond = True
                    else:
                        cond = False
                stack.append((cond, cond))
            elif kw == "elif":
                taking, any_taken = stack.pop()
                stack.append((False, any_taken))
            elif kw == "else":
                taking, any_taken = stack.pop()
                stack.append((not any_taken, True))
            else:
                stack.pop()
            continue
        if all(t for t, _ in stack):
            out.append(line)
    return "\n".join(out)


def strip_comments(text):
    text = re.sub(r"/\*.*?\*/", " ", text, flags=re.S)
    return re.sub(r"//[^\n]*", " ", text)


def table(text, name):
    """The initialiser list of `name[...] = { ... };` as a list of tokens (numbers, braces dropped)."""
    m = re.search(r"\b%s\s*\[[^\]]*\]\s*(\[[^\]]*\]\s*)*=\s*\{" % re.escape(name), text)
    if not m:
        raise SystemExit("table %s not found" % name)
    depth, i = 1, m.end()
    while depth:
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
        i += 1
    body = text[m.end():i - 1].replace("{", " ").replace("}", " ")
    return [t.strip() for t in body.split(",") if t.strip()]


def float_bits(literal):
    """The bits of the float nearest to the C literal (correct rounding, ties to even, as C compilers do)."""
    lit = literal.rstrip("fF")
    exact = Fraction(lit)
    d = float(exact)
    f = struct.unpack("<f", struct.pack("<f", d))[0]
    best = None
    bits0 = struct.unpack("<I", struct.pack("<f", f))[0]
    for delta in (-1, 0, 1):
        b = bits0 + delta if f >= 0 else bits0 - delta
        if b < 0 or b > 0xFFFFFFFF:
            continue
        v = struct.unpack("<f", struct.pack("<I", b))[0]
        err = abs(Fraction(v) - exact)
        key = (err, b & 1)
        if best is None or key < best[0]:
            best = (key, b)
    return best[1]


def int_value(tok):
    tok = tok.strip().rstrip("uUlL")
    return int(tok, 0)


def vb_lines(prefix, items, per_line):
    """'s = s & "...."' lines of concatenated items"""
    lines = []
    for i in range(0, len(items), per_line):
        lines.append('    %s = %s & "%s"' % (prefix, prefix, "".join(items[i:i + per_line])))
    return lines


def main():
    with open(os.path.join(CELT, "static_modes_float.h"), "r", encoding="latin-1") as f:
        modes_h = strip_comments(preprocess(f.read()))
    with open(os.path.join(CELT, "modes.c"), "r", encoding="latin-1") as f:
        modes_c = strip_comments(preprocess(f.read()))
    with open(os.path.join(CELT, "cwrs.c"), "r", encoding="latin-1") as f:
        cwrs_c = strip_comments(preprocess(f.read()))
    with open(os.path.join(CELT, "quant_bands.c"), "r", encoding="latin-1") as f:
        qb_c = strip_comments(preprocess(f.read()))

    floats = {}     # name -> list of bits
    ints = {}       # name -> (vbtype, values)
    floats["window120"] = [float_bits(t) for t in table(modes_h, "window120")]
    ints["logN400"] = ("Integer", [int_value(t) for t in table(modes_h, "logN400")])
    ints["cache_index50"] = ("Integer", [int_value(t) for t in table(modes_h, "cache_index50")])
    ints["cache_bits50"] = ("Byte", [int_value(t) for t in table(modes_h, "cache_bits50")])
    ints["cache_caps50"] = ("Byte", [int_value(t) for t in table(modes_h, "cache_caps50")])
    tw = table(modes_h, "fft_twiddles48000_960")
    floats["fft_twiddles48000_960_r"] = [float_bits(t) for t in tw[0::2]]
    floats["fft_twiddles48000_960_i"] = [float_bits(t) for t in tw[1::2]]
    for n in (480, 240, 120, 60):
        ints["fft_bitrev%d" % n] = ("Integer", [int_value(t) for t in table(modes_h, "fft_bitrev%d" % n)])
    floats["mdct_twiddles960"] = [float_bits(t) for t in table(modes_h, "mdct_twiddles960")]
    ints["eband5ms"] = ("Integer", [int_value(t) for t in table(modes_c, "eband5ms")])
    ints["band_allocation"] = ("Byte", [int_value(t) for t in table(modes_c, "band_allocation")])
    ints["CELT_PVQ_U_DATA"] = ("Double", [int_value(t) for t in table(cwrs_c, "CELT_PVQ_U_DATA")])
    ints["e_prob_model"] = ("Byte", [int_value(t) for t in table(qb_c, "e_prob_model")])
    # eMeans: the float build's table (the FIXED_POINT one is preprocessed away)
    floats["eMeans"] = [float_bits(t) for t in table(qb_c, "eMeans")]

    expect = {"window120": 120, "logN400": 21, "cache_index50": 105, "cache_bits50": 392, "cache_caps50": 168,
              "fft_twiddles48000_960_r": 480, "fft_twiddles48000_960_i": 480, "fft_bitrev480": 480,
              "fft_bitrev240": 240, "fft_bitrev120": 120, "fft_bitrev60": 60, "mdct_twiddles960": 1800,
              "eband5ms": 22, "band_allocation": 231, "CELT_PVQ_U_DATA": 1272, "e_prob_model": 336, "eMeans": 25}
    for k, n in expect.items():
        got = len(floats[k]) if k in floats else len(ints[k][1])
        if got != n:
            raise SystemExit("%s: %d values, expected %d" % (k, got, n))

    src = {"window120": "static_modes_float.h", "logN400": "static_modes_float.h",
           "cache_index50": "static_modes_float.h", "cache_bits50": "static_modes_float.h",
           "cache_caps50": "static_modes_float.h", "fft_twiddles48000_960_r": "static_modes_float.h (.r)",
           "fft_twiddles48000_960_i": "static_modes_float.h (.i)", "fft_bitrev480": "static_modes_float.h",
           "fft_bitrev240": "static_modes_float.h", "fft_bitrev120": "static_modes_float.h",
           "fft_bitrev60": "static_modes_float.h", "mdct_twiddles960": "static_modes_float.h",
           "eband5ms": "modes.c", "band_allocation": "modes.c", "CELT_PVQ_U_DATA": "cwrs.c (uint32 as Double)",
           "e_prob_model": "quant_bands.c ([4][2][42] flattened)", "eMeans": "quant_bands.c (float build)"}
    order = ["eband5ms", "band_allocation", "window120", "logN400", "cache_index50", "cache_bits50",
             "cache_caps50", "fft_twiddles48000_960_r", "fft_twiddles48000_960_i", "fft_bitrev480",
             "fft_bitrev240", "fft_bitrev120", "fft_bitrev60", "mdct_twiddles960", "CELT_PVQ_U_DATA",
             "e_prob_model", "eMeans"]

    out = []
    out.append('Attribute VB_Name = "celt_tables_vb"')
    out.append("' (VB) GENERATED by tools/gen_celt_tables.py from libopus 1.5.2 (ref/opus/celt): never edit. The constant")
    out.append("' tables of the CELT decoder: celt/static_modes_float.h (the static 48 kHz / 960 mode of the float build),")
    out.append("' celt/modes.c (eband5ms, band_allocation), celt/cwrs.c (CELT_PVQ_U_DATA, CUSTOM_MODES undefined),")
    out.append("' celt/quant_bands.c (e_prob_model, eMeans). Floats are the bits of C's float literals, so the tables equal")
    out.append("' the C ones bit for bit; celt_tables_init fills them once (VB has no array initialisers).")
    out.append("Option Explicit")
    out.append("")
    out.append('Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)')
    out.append("")
    for k in order:
        t = "Single" if k in floats else ints[k][0]
        n = len(floats[k]) if k in floats else len(ints[k][1])
        out.append("Public %s() As %s%s' %d: %s" % (k, t, " " * max(1, 30 - len(k) - len(t)), n, src[k]))
    out.append("")
    out.append("Private celt_tables_done As Boolean")
    out.append("")
    out.append("' (VB) fills the tables (once)")
    out.append("Public Sub celt_tables_init()")
    out.append("    If celt_tables_done Then Exit Sub")
    for k in order:
        out.append("    init_%s" % k)
    out.append("    celt_tables_done = True")
    out.append("End Sub")
    out.append("")
    out.append("' (VB) a hex digit's value")
    out.append("Private Function hexval(ByVal ch As Long) As Long")
    out.append("    If ch <= 57 Then hexval = ch - 48 Else hexval = ch - 55")
    out.append("End Function")
    out.append("")
    out.append("' (VB) the 8-digit hex numbers of s (the bits of each value) into a(0...)")
    out.append("Private Sub hex_singles(s As String, a() As Single)")
    out.append("    Dim i As Long, j As Long, v As Long, b() As Byte, n As Long")
    out.append("    b = StrConv(s, vbFromUnicode)")
    out.append("    n = (UBound(b) + 1) \\ 8")
    out.append("    ReDim a(0 To n - 1)")
    out.append("    For i = 0 To n - 1")
    out.append("        v = hexval(b(i * 8)) And 7")
    out.append("        For j = 1 To 7")
    out.append("            v = v * 16 + hexval(b(i * 8 + j))")
    out.append("        Next")
    out.append("        If hexval(b(i * 8)) >= 8 Then v = v Or &H80000000")
    out.append("        CopyMemory a(i), v, 4")
    out.append("    Next")
    out.append("End Sub")
    out.append("")
    out.append("' (VB) the comma-separated decimal numbers of s")
    out.append("Private Function dec_values(s As String) As String()")
    out.append("    dec_values = Split(s, \",\")")
    out.append("End Function")
    out.append("")
    for k in order:
        out.append("Private Sub init_%s()" % k)
        out.append("    Dim s As String, v() As String, i As Long")
        if k in floats:
            items = ["%08X" % b for b in floats[k]]
            out.extend(vb_lines("s", items, 100))
            out.append("    hex_singles s, %s" % k)
        else:
            t, vals = ints[k]
            items = ["%d," % v for v in vals]
            items[-1] = items[-1][:-1]
            per = 16 if t == "Double" else 40
            out.extend(vb_lines("s", items, per))
            out.append("    v = dec_values(s)")
            out.append("    ReDim %s(0 To UBound(v))" % k)
            out.append("    For i = 0 To UBound(v)")
            out.append("        %s(i) = %s(v(i))" % (k, {"Integer": "CInt", "Byte": "CByte", "Double": "CDbl"}[t]))
            out.append("    Next")
        out.append("End Sub")
        out.append("")
    data = ("\r\n".join(out) + "\r\n").encode("ascii")
    for line in data.split(b"\r\n"):
        if len(line) > 1000:
            raise SystemExit("line too long: %d" % len(line))
    with open(OUT, "wb") as f:
        f.write(data)
    print("wrote %s (%d lines)" % (os.path.relpath(OUT, ROOT), len(out)))


if __name__ == "__main__":
    sys.exit(main())
