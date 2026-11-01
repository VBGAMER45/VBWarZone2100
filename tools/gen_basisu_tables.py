"""
gen_basisu_tables.py - writes src/basis/basisu_tables_vb.bas (cluster T, fork T2): the constant tables of Basis
Universal's transcoder (ref/basis_universal @ df079ec, transcoder/basisu_transcoder.cpp) that its UASTC LDR paths
read (unpack_uastc, UASTC -> ASTC / BC7 / BC1 / BC3 / BC4 / BC5), taken from the C text (VB6 has no array
initialisers). A C array of structs becomes one VB array per field (<name>_<field>); a C array [a][b] is a VB
array (0 To a - 1, 0 To b - 1). The tables uastc_init and basisu_transcoder_init compute (g_astc_unquant, the BC7
optimal endpoints, the BC1 single colour tables) are basisu_uastc_c's.

  python -I tools\\gen_basisu_tables.py [--stage DIR]     (default: writes into src\\basis)
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CPP = os.path.join(ROOT, "ref", "basis_universal", "transcoder", "basisu_transcoder.cpp")

# (C name, dims, VB type, struct fields or None)
TABLES = [
    ("g_astc_bc7_common_partitions2", [30], "Long", ["m_bc7", "m_astc", "m_invert"]),
    ("g_bc7_3_astc2_common_partitions", [19], "Long", ["m_bc73", "m_astc2", "k"]),
    ("g_astc_bc7_common_partitions3", [11], "Long", ["m_bc7", "m_astc", "m_astc_to_bc7_perm"]),
    ("g_astc_to_bc7_partition_index_perm_tables", [6, 3], "Long", None),
    ("g_astc_bc7_patterns2", [30, 16], "Long", None),
    ("g_astc_bc7_patterns3", [11, 16], "Long", None),
    ("g_bc7_3_astc2_patterns2", [19, 16], "Long", None),
    ("g_astc_bc7_pattern2_anchors", [30, 3], "Long", None),
    ("g_astc_bc7_pattern3_anchors", [11, 3], "Long", None),
    ("g_bc7_3_astc2_patterns2_anchors", [19, 3], "Long", None),
    ("g_uastc_mode_huff_codes", [20, 2], "Long", None),
    ("g_uastc_huff_modes", [128], "Long", None),
    ("g_uastc_mode_weight_bits", [19], "Long", None),
    ("g_uastc_mode_weight_ranges", [19], "Long", None),
    ("g_uastc_mode_endpoint_ranges", [19], "Long", None),
    ("g_uastc_mode_subsets", [19], "Long", None),
    ("g_uastc_mode_planes", [19], "Long", None),
    ("g_uastc_mode_comps", [19], "Long", None),
    ("g_uastc_mode_has_etc1_bias", [19], "Long", None),
    ("g_uastc_mode_has_bc1_hint0", [19], "Long", None),
    ("g_uastc_mode_has_bc1_hint1", [19], "Long", None),
    ("g_uastc_mode_cem", [19], "Long", None),
    ("g_uastc_mode_has_alpha", [19], "Long", None),
    ("g_uastc_mode_is_la", [19], "Long", None),
    ("g_uastc_mode_total_hint_bits", [19], "Long", None),
    ("g_astc_bise_range_table", [21, 3], "Long", None),
    ("g_bc7_weights1", [2], "Long", None),
    ("g_bc7_weights2", [4], "Long", None),
    ("g_bc7_weights3", [8], "Long", None),
    ("g_bc7_weights4", [16], "Long", None),
    ("g_astc_weights4", [16], "Long", None),
    ("g_astc_weights5", [32], "Long", None),
    ("g_astc_weights_3levels", [3], "Long", None),
    ("g_bc7_partition1", [16], "Long", None),
    ("g_bc7_partition2", [64 * 16], "Long", None),
    ("g_bc7_partition3", [64 * 16], "Long", None),
    ("g_bc7_table_anchor_index_second_subset", [64], "Long", None),
    ("g_bc7_table_anchor_index_third_subset_1", [64], "Long", None),
    ("g_bc7_table_anchor_index_third_subset_2", [64], "Long", None),
    ("g_bc7_num_subsets", [8], "Long", None),
    ("g_bc7_partition_bits", [8], "Long", None),
    ("g_bc7_color_index_bitcount", [8], "Long", None),
    ("g_bc7_mode_has_p_bits", [8], "Long", None),
    ("g_bc7_mode_has_shared_p_bits", [8], "Long", None),
    ("g_bc7_color_precision_table", [8], "Long", None),
    ("g_bc7_alpha_precision_table", [8], "Long", None),
    ("g_bc7_alpha_index_bitcount", [8], "Long", None),
    ("g_astc_trit_encode", [243], "Long", None),
    ("g_astc_quint_encode", [125], "Long", None),
    ("g_uastc_mode_astc_block_mode", [19], "Long", None),
    ("s_uastc5_to_bc1", [32], "Long", None),
    ("s_uastc4_to_bc1", [16], "Long", None),
    ("s_uastc3_to_bc1", [8], "Long", None),
    ("s_uastc2_to_bc1", [4], "Long", None),
    ("s_uastc1_to_bc1", [2], "Long", None),
]


def initializer(text, name):
    """the text between the braces of `name[...] = { ... };` (the first definition with an initialiser)"""
    m = re.search(r"\b" + re.escape(name) + r"\s*(\[[^\]]*\]\s*)+=\s*\{", text)
    if not m:
        sys.exit("gen_basisu_tables: %s not found" % name)
    i = m.end()
    depth = 1
    j = i
    while depth:
        c = text[j]
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        j += 1
    return text[i:j - 1]


def values(body):
    body = re.sub(r"//[^\n]*", "", body)
    out = []
    for tok in re.findall(r"0[xX][0-9a-fA-F]+|\d+|true|false", body):
        if tok == "true":
            out.append(1)
        elif tok == "false":
            out.append(0)
        else:
            out.append(int(tok, 0))
    return out


def chunks(vals, per_line=40):
    for k in range(0, len(vals), per_line):
        yield ",".join(str(v) for v in vals[k:k + per_line])


def main():
    out_dir = os.path.join(ROOT, "src", "basis")
    if "--stage" in sys.argv:
        out_dir = os.path.join(sys.argv[sys.argv.index("--stage") + 1], "basis")
    text = open(CPP, encoding="utf-8").read()
    decls, inits, subs = [], [], []
    for name, dims, vbtype, fields in TABLES:
        body_text = initializer(text, name)
        rows = re.findall(r"\{([^{}]*)\}", body_text)
        if rows:
            # a row per inner brace group, padded with zeros as C pads a short initialiser
            width = len(fields) if fields else dims[-1]
            vals = []
            for r in rows:
                rv = values(r)
                vals += rv + [0] * (width - len(rv))
        else:
            vals = values(body_text)
        total = 1
        for d in dims:
            total *= d
        nf = len(fields) if fields else 1
        if len(vals) != total * nf:
            sys.exit("gen_basisu_tables: %s has %d values, expected %d" % (name, len(vals), total * nf))
        bounds = ", ".join("0 To %d" % (d - 1) for d in dims)
        names = ["%s_%s" % (name, f) for f in fields] if fields else [name]
        for n in names:
            decls.append("Public %s(%s) As %s" % (n, bounds, vbtype))
        inits.append("    init_%s" % name)
        body = ["Private Sub init_%s()" % name, "    Dim s As String, v() As Long, i As Long"]
        for c in chunks(vals):
            body.append('    s = s & "%s,"' % c)
        body.append("    dec_values s, v")
        if len(dims) == 1:
            idx = "(i)"
            loop = ["    For i = 0 To %d" % (total - 1)]
        else:
            idx = "(i \\ %d, i Mod %d)" % (dims[1], dims[1])
            loop = ["    For i = 0 To %d" % (total - 1)]
        body += loop
        for f, n in enumerate(names):
            body.append("        %s%s = v(i * %d + %d)" % (n, idx, nf, f))
        body += ["    Next", "End Sub", ""]
        subs.append("\n".join(body))
    lines = [
        'Attribute VB_Name = "basisu_tables_vb"',
        "' (VB) GENERATED by tools/gen_basisu_tables.py from Basis Universal's transcoder (ref/basis_universal @ df079ec,",
        "' transcoder/basisu_transcoder.cpp): never edit. The constant tables its UASTC LDR paths read (unpack_uastc and",
        "' the UASTC -> ASTC, BC7, BC1, BC3, BC4, BC5 transcoders). A C array of structs is one array per field",
        "' (<name>_<field>); a C [a][b] array is (0 To a - 1, 0 To b - 1); uint8_t / bool / int values are Longs.",
        "' basisu_tables_init fills them once (VB has no array initialisers).",
        "Option Explicit",
        "",
    ] + decls + [
        "",
        "Private tables_done As Boolean",
        "",
        "' (VB) fills the tables (once)",
        "Public Sub basisu_tables_init()",
        "    If tables_done Then Exit Sub",
    ] + inits + [
        "    tables_done = True",
        "End Sub",
        "",
        "' (VB) the comma-separated decimal values of s (a trailing comma) into v (0-based)",
        "Private Sub dec_values(ByVal s As String, v() As Long)",
        "    Dim parts() As String, i As Long",
        "    parts = Split(s, \",\")",
        "    ReDim v(0 To UBound(parts) - 1)",
        "    For i = 0 To UBound(parts) - 1",
        "        v(i) = CLng(parts(i))",
        "    Next",
        "End Sub",
        "",
    ] + subs
    data = "\n".join(lines).rstrip("\n").replace("\r\n", "\n").replace("\n", "\r\n") + "\r\n"
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "basisu_tables_vb.bas")
    with open(path, "wb") as f:
        f.write(data.encode("ascii"))
    print("wrote %s (%d tables)" % (path, len(TABLES)))


if __name__ == "__main__":
    main()
