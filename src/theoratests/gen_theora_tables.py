"""
gen_theora_tables.py - writes src/sequence/theora_tables_vb.bas: libtheora 1.2.0's constant tables (the
decoder's) as VB arrays, taken from the C text of ref/theora/lib (the initialisers are evaluated with the
headers' macros), and an init procedure that fills them (VB has no array initialisers).

  python -I src\\theoratests\\gen_theora_tables.py

Multi-dimensional C tables become flat 0-based Long arrays in the C's (row-major) order: T[i][j] of a
[n][m] table is T(i * m + j). Never edit the output; run this again.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LIB = os.path.join(ROOT, "ref", "theora", "lib")
OUT = os.path.join(ROOT, "src", "sequence", "theora_tables_vb.bas")

# (C file, C name, VB name, comment); static tables inside functions are found the same way
TABLES = [
    ("internal.c", "OC_FZIG_ZAG", "OC_FZIG_ZAG", "internal.c: zig zag index -> coefficient number (64 past the end)"),
    ("internal.c", "OC_IZIG_ZAG", "OC_IZIG_ZAG", "internal.c: coefficient number -> zig zag index"),
    ("internal.c", "OC_MB_MAP", "OC_MB_MAP", "internal.c: [2][2] physical -> bitstream macro block order"),
    ("internal.c", "OC_MB_MAP_IDXS", "OC_MB_MAP_IDXS", "internal.c: [TH_PF_NFORMATS][12] valid oc_mb_map indices"),
    ("internal.c", "OC_MB_MAP_NIDXS", "OC_MB_MAP_NIDXS", "internal.c: [TH_PF_NFORMATS] their numbers"),
    ("internal.c", "OC_DCT_TOKEN_EXTRA_BITS", "OC_DCT_TOKEN_EXTRA_BITS", "internal.c: extra bits per DCT token"),
    ("huffdec.c", "OC_DCT_TOKEN_MAP", "OC_DCT_TOKEN_MAP", "huffdec.c: spec token -> internal token"),
    ("huffdec.c", "OC_DCT_TOKEN_MAP_LOG_NENTRIES", "OC_DCT_TOKEN_MAP_LOG_NENTRIES",
     "huffdec.c: log2 of the internal tokens per spec token"),
    ("decode.c", "OC_MODE_ALPHABETS", "OC_MODE_ALPHABETS", "decode.c: [7][OC_NMODES] mode alphabets"),
    ("decode.c", "OC_INTERNAL_DCT_TOKEN_EXTRA_BITS", "OC_INTERNAL_DCT_TOKEN_EXTRA_BITS",
     "decode.c: extra bits per internal DCT token"),
    ("decode.c", "OC_DCT_CODE_WORD", "OC_DCT_CODE_WORD", "decode.c: the code word of each internal token"),
    ("decode.c", "OC_SB_RUN_TREE", "OC_SB_RUN_TREE", "decode.c (oc_sb_run_unpack): the super block run tree"),
    ("decode.c", "OC_BLOCK_RUN_TREE", "OC_BLOCK_RUN_TREE", "decode.c (oc_block_run_unpack): the block run tree"),
    ("decode.c", "OC_VLC_MODE_TREE", "OC_VLC_MODE_TREE", "decode.c: the variable length mode tree"),
    ("decode.c", "OC_CLC_MODE_TREE", "OC_CLC_MODE_TREE", "decode.c: the constant length mode tree"),
    ("decode.c", "OC_VLC_MV_COMP_TREE", "OC_VLC_MV_COMP_TREE", "decode.c: the variable length MV tree"),
    ("decode.c", "OC_CLC_MV_COMP_TREE", "OC_CLC_MV_COMP_TREE", "decode.c: the constant length MV tree"),
    ("decode.c", "OC_HUFF_LIST_MAX", "OC_HUFF_LIST_MAX",
     "decode.c (oc_dec_residual_tokens_unpack): the end of each Huffman table group"),
    ("decode.c", "OC_MOD_MAX", "OC_MOD_MAX", "decode.c (oc_dering_block)"),
    ("decode.c", "OC_MOD_SHIFT", "OC_MOD_SHIFT", "decode.c (oc_dering_block)"),
    ("state.c", "SB_MAP", "OC_SB_MAP", "state.c (oc_sb_create_plane_mapping): SB_MAP[4][4][2], the Hilbert curve"),
    ("state.c", "OC_MVMAP", "OC_MVMAP", "state.c (oc_state_get_mv_offsets): [2][64]"),
    ("state.c", "OC_MVMAP2", "OC_MVMAP2", "state.c (oc_state_get_mv_offsets): [2][64]"),
    ("quant.c", "OC_DC_QUANT_MIN", "OC_DC_QUANT_MIN", "quant.c: [2]"),
    ("quant.c", "OC_AC_QUANT_MIN", "OC_AC_QUANT_MIN", "quant.c: [2]"),
]

HEADERS = ["state.h", "huffman.h", "decint.h", "internal.h", "quant.h"]


def strip_comments(text):
    text = re.sub(r"/\*.*?\*/", " ", text, flags=re.S)
    return re.sub(r"//[^\n]*", " ", text)


def macros():
    """Simple numeric #defines of the headers and decode.c, evaluated."""
    env = {}
    for name in HEADERS + ["decode.c"]:
        text = strip_comments(open(os.path.join(LIB, name), encoding="latin-1").read())
        for m in re.finditer(r"^\s*#\s*define\s+(\w+)\s+(\([^\n]*\)|-?\d+)\s*$", text, flags=re.M):
            key, val = m.group(1), m.group(2)
            try:
                env[key] = eval(val, {}, env)
            except Exception:
                pass
    eb, rl, fb, ms = env["OC_DCT_CW_EOB_SHIFT"], env["OC_DCT_CW_RLEN_SHIFT"], env["OC_DCT_CW_FLIP_BIT"], \
        env["OC_DCT_CW_MAG_SHIFT"]
    env["OC_DCT_CW_PACK"] = lambda e, r, m, f: (e << eb) | (r << rl) | (f << fb) | ((m - f) * (1 << ms))
    env["OC_DCT_CW_FINISH"] = 0
    env["TH_NDCT_TOKENS"] = 32
    env["TH_PF_NFORMATS"] = 4
    return env


def table_values(cfile, cname, env):
    text = strip_comments(open(os.path.join(LIB, cfile), encoding="latin-1").read())
    m = re.search(r"\b%s\s*((?:\[[^\]]*\])+)\s*=\s*\{" % re.escape(cname), text)
    if not m:
        raise SystemExit("%s: %s not found" % (cfile, cname))
    start = m.end() - 1
    depth = 0
    for i in range(start, len(text)):
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
            if depth == 0:
                body = text[start + 1:i]
                break
    dims = [eval(d.strip("[]") or "0", {}, env) for d in re.findall(r"\[[^\]]*\]", m.group(1))]

    def split_top(s):
        """Items separated by commas outside parentheses and braces."""
        items, depth, cur = [], 0, ""
        for ch in s:
            if ch in "({":
                depth += 1
            elif ch in ")}":
                depth -= 1
            if ch == "," and depth == 0:
                items.append(cur)
                cur = ""
            else:
                cur += ch
        items.append(cur)
        return [i.strip() for i in items if i.strip()]

    def flatten(s, level):
        """The values of an initialiser at dimension level, each sub-brace zero-filled to its size (C)."""
        out = []
        for item in split_top(s):
            if item.startswith("{"):
                sub = flatten(item[1:-1], level + 1)
                size = 1
                for d in dims[level + 1:]:
                    size *= d
                out += sub + [0] * (size - len(sub))
            else:
                # C casts are not in these initialisers; no division is used
                out.append(int(eval(item, {}, env)))
        return out

    values = flatten(body, 0)
    size = 1
    for d in dims:
        size *= d
    if size and len(values) < size:
        values += [0] * (size - len(values))   # C zero-fills the rest
    return values, dims


def vb_lines(vbname, values):
    out = []
    text = ",".join(str(v) for v in values)
    chunks = []
    while text:
        if len(text) <= 96:
            chunks.append(text)
            break
        cut = text.rfind(",", 0, 96) + 1
        chunks.append(text[:cut])
        text = text[cut:]
    out.append('    s = "%s"' % chunks[0])
    for c in chunks[1:]:
        out.append('    s = s & "%s"' % c)
    out.append("    oc_table_fill %s, %d, s" % (vbname, len(values)))
    return out


def main():
    env = macros()
    decls = []
    body = []
    for cfile, cname, vbname, comment in TABLES:
        values, dims = table_values(cfile, cname, env)
        decls.append("Public %s() As Long%s' %s, %d entries" % (vbname, " " * max(1, 34 - len(vbname)), comment,
                                                                  len(values)))
        body += vb_lines(vbname, values)
    lines = [
        'Attribute VB_Name = "theora_tables_vb"',
        "' Generated by src/theoratests/gen_theora_tables.py from ref/theora/lib (libtheora 1.2.0): internal.c,",
        "' huffdec.c, decode.c, state.c, quant.c. Do not edit; run the script again.",
        "' The decoder's constant tables as flat 0-based Long arrays in the C's row-major order (T[i][j] of a [n][m]",
        "' table is T(i * m + j)); theora_tables_init fills them once (the decoder's entry points call it).",
        "Option Explicit",
        "",
    ] + decls + [
        "",
        "Private tablesReady As Boolean",
        "",
        "Private Sub oc_table_fill(t() As Long, ByVal n As Long, ByVal s As String)",
        "    Dim parts() As String, i As Long",
        "    parts = Split(s, \",\")",
        "    ReDim t(0 To n - 1)",
        "    For i = 0 To n - 1",
        "        t(i) = CLng(parts(i))",
        "    Next",
        "End Sub",
        "",
        "Public Sub theora_tables_init()",
        "    Dim s As String",
        "    If tablesReady Then Exit Sub",
    ] + body + [
        "    tablesReady = True",
        "End Sub",
        "",
    ]
    data = ("\r\n".join(lines)).encode("ascii")
    with open(OUT, "wb") as f:
        f.write(data)
    print("%s: %d tables" % (os.path.relpath(OUT, ROOT), len(TABLES)))


if __name__ == "__main__":
    main()
