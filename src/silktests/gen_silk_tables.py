"""
gen_silk_tables.py - writes src/sound/silk_tables_c.bas from libopus 1.5.2's SILK table sources (ref/opus/silk:
tables_*.c, table_LSF_cos.c, pitch_est_tables.c, resampler_rom.c and resampler_rom.h's static tables).

  python -I src\\silktests\\gen_silk_tables.py

VB has no array initialisers, so each C element type gets one ROM array that silk_tables_init fills once from
decimal strings: opus_uint8 tables live in silk_rom_u8 (Byte), opus_int8 in silk_rom_i8 and opus_int16 in
silk_rom_i16 (Integer), opus_int32 in silk_rom_i32 (Long). A table's C name is a Public Const holding its offset
in its ROM, so a C pointer to a table (or into one) is an offset, and the C's tables of pointers
(silk_LBRR_flags_iCDF_ptr, silk_LTP_gain_iCDF_ptrs, ...) are Long arrays of offsets. The two NLSF codebooks
(silk_NLSF_CB_struct) are silk_NLSF_CBs(1) and (2); silk_NLSF_CB_NB_MB and silk_NLSF_CB_WB are their indices.
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
SILK = os.path.join(ROOT, "ref", "opus", "silk")
OUT = os.path.join(ROOT, "src", "sound", "silk_tables_c.bas")

FILES = ["tables_gain.c", "tables_LTP.c", "tables_NLSF_CB_NB_MB.c", "tables_NLSF_CB_WB.c", "tables_other.c",
         "tables_pitch_lag.c", "tables_pulses_per_block.c", "table_LSF_cos.c", "pitch_est_tables.c",
         "resampler_rom.h", "resampler_rom.c"]

# the macros the tables use (define.h, pitch_est_defines.h, resampler_rom.h)
MACROS = {
    "N_LEVELS_QGAIN": 64, "MAX_DELTA_GAIN_QUANT": 36, "MIN_DELTA_GAIN_QUANT": -4, "PITCH_EST_MAX_LAG_MS": 18,
    "PITCH_EST_MIN_LAG_MS": 2, "N_RATE_LEVELS": 10, "SILK_MAX_PULSES": 16, "NB_LTP_CBKS": 3,
    "STEREO_QUANT_TAB_SIZE": 16, "TRANSITION_INT_NUM": 5, "TRANSITION_NB": 3, "TRANSITION_NA": 2,
    "LSF_COS_TAB_SZ_FIX": 128, "OFFSET_VL_Q10": 32, "OFFSET_VH_Q10": 100, "OFFSET_UVL_Q10": 100,
    "OFFSET_UVH_Q10": 240, "PE_MAX_NB_SUBFR": 4, "PE_NB_CBKS_STAGE2_EXT": 11, "PE_NB_CBKS_STAGE3_MAX": 34,
    "PE_NB_CBKS_STAGE2_10MS": 3, "PE_NB_CBKS_STAGE3_MID": 24, "PE_NB_CBKS_STAGE3_MIN": 16, "PE_NB_CBKS_STAGE3_10MS": 12, "SILK_PE_MAX_COMPLEX": 2,
    "RESAMPLER_DOWN_ORDER_FIR0": 18, "RESAMPLER_DOWN_ORDER_FIR1": 24, "RESAMPLER_DOWN_ORDER_FIR2": 36,
    "RESAMPLER_ORDER_FIR_12": 8,
}

ROMS = {"opus_uint8": "u8", "opus_int8": "i8", "opus_int16": "i16", "opus_int32": "i32"}
VBTYPE = {"u8": "Byte", "i8": "Integer", "i16": "Integer", "i32": "Long"}
RANGE = {"u8": (0, 255), "i8": (-128, 127), "i16": (-32768, 32767), "i32": (-2 ** 31, 2 ** 31 - 1)}

ARRAY = re.compile(r"(?:static\s+)?(?:silk_DWORD_ALIGN\s+)?const\s+(opus_uint8|opus_int8|opus_int16|opus_int32)\s+"
                   r"(\w+)\s*((?:\[[^\]]*\]\s*)+)=\s*\{(.*?)\}\s*;", re.S)
PTRS = re.compile(r"const\s+(opus_uint8|opus_int8)\s*\*\s*const\s+(\w+)\s*\[[^\]]*\]\s*=\s*\{(.*?)\}\s*;", re.S)
STRUCT = re.compile(r"const\s+silk_NLSF_CB_struct\s+(\w+)\s*=\s*\{(.*?)\}\s*;", re.S)


def strip_comments(text):
    text = re.sub(r"/\*.*?\*/", " ", text, flags=re.S)
    return re.sub(r"//[^\n]*", " ", text)


def c_eval(expr):
    """an integer constant expression of the tables (macros, SILK_FIX_CONST, arithmetic)"""
    e = expr.strip()

    def fix_const(m):
        c, q = m.group(1), m.group(2)
        # SILK_FIX_CONST( C, Q ): ((opus_int32)((C) * ((opus_int64)1 << (Q)) + 0.5)), C in double
        return str(int(eval(c, {}, {}) * (1 << int(eval(q, {}, MACROS))) + 0.5))
    e = re.sub(r"SILK_FIX_CONST\s*\(\s*([^,]+),\s*([^)]+)\)", fix_const, e)
    for k, v in MACROS.items():
        e = re.sub(r"\b%s\b" % k, "(%d)" % v, e)
    e = e.replace("/", "//")
    return int(eval(e, {}, {}))


def main():
    tables = []          # (name, rom, values, dims text, file)
    ptrs = []            # (name, rom, [target names])
    structs = []         # (name, [fields])
    for fn in FILES:
        with open(os.path.join(SILK, fn), "r", encoding="latin-1") as f:
            text = strip_comments(f.read())
        for m in ARRAY.finditer(text):
            ctype, name, dims, body = m.groups()
            vals = [c_eval(v) for v in body.replace("{", " ").replace("}", " ").split(",") if v.strip()]
            size = 1
            for d in re.findall(r"\[([^\]]*)\]", dims):
                size *= c_eval(d)
            if size != len(vals):
                sys.exit("%s: %d values for %d elements" % (name, len(vals), size))
            rom = ROMS[ctype]
            lo, hi = RANGE[rom]
            for v in vals:
                if not lo <= v <= hi:
                    sys.exit("%s: %d out of range" % (name, v))
            tables.append((name, rom, vals, " ".join(dims.split()), fn))
        for m in PTRS.finditer(text):
            ctype, name, body = m.groups()
            targets = []
            for item in body.split(","):
                item = item.strip()
                if not item:
                    continue
                t = re.search(r"(silk_\w+)", item).group(1)
                targets.append(t)
            ptrs.append((name, ROMS[ctype], targets))
        for m in STRUCT.finditer(text):
            name, body = m.groups()
            fields, depth, part = [], 0, ""
            for ch in body + ",":
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                if ch == "," and depth == 0:
                    if part.strip():
                        fields.append(part.strip())
                    part = ""
                else:
                    part += ch
            structs.append((name, fields))

    offsets = {}
    rom_vals = {"u8": [], "i8": [], "i16": [], "i32": []}
    for name, rom, vals, dims, fn in tables:
        offsets[name] = (rom, len(rom_vals[rom]))
        rom_vals[rom].extend(vals)

    L = []
    L.append('Attribute VB_Name = "silk_tables_c"')
    L.append("' libopus 1.5.2 (ref/opus/silk): the SILK tables of " + ", ".join(FILES[:5]) + ",")
    L.append("' " + ", ".join(FILES[5:]) + " (its static tables).")
    L.append("' GENERATED by src/silktests/gen_silk_tables.py from the C text: do not edit.")
    L.append("' (VB) each C element type has one ROM array, filled once by silk_tables_init: opus_uint8 tables in")
    L.append("' silk_rom_u8 (Byte), opus_int8 in silk_rom_i8 and opus_int16 in silk_rom_i16 (Integer), opus_int32 in")
    L.append("' silk_rom_i32 (Long). A table's C name is a Const: its offset in its ROM (a 2-D table row by row, as C")
    L.append("' lays it out), so a C pointer to a table is an offset. A table of pointers is a Long array of offsets.")
    L.append("' The NLSF codebooks (const silk_NLSF_CB_struct) are silk_NLSF_CBs(1 To 2); silk_NLSF_CB_NB_MB and")
    L.append("' silk_NLSF_CB_WB are their indices (a const silk_NLSF_CB_struct * is such an index).")
    L.append("Option Explicit")
    L.append("")
    for rom in ("u8", "i8", "i16", "i32"):
        L.append("Public silk_rom_%s() As %s" % (rom, VBTYPE[rom]))
    L.append("")
    for name, rom, vals, dims, fn in tables:
        L.append("Public Const %s As Long = %d%s' %s: %s%s (%d)" % (
            name, offsets[name][1], " " * max(1, 8 - len(str(offsets[name][1]))), fn,
            [k for k, v in ROMS.items() if v == rom][0], dims, len(vals)))
    L.append("")
    L.append("' tables of pointers: offsets")
    for name, rom, targets in ptrs:
        L.append("Public %s(0 To %d) As Long%s' into silk_rom_%s" % (name, len(targets) - 1, " " * max(1, 20 - len(name)), rom))
    L.append("")
    L.append("' const silk_NLSF_CB_struct *")
    for i, (name, fields) in enumerate(structs):
        L.append("Public Const %s As Long = %d" % (name, i + 1))
    L.append("Public silk_NLSF_CBs(1 To %d) As silk_NLSF_CB_struct" % len(structs))
    L.append("")
    L.append("Private silk_tables_ready As Boolean")
    L.append("")
    # fill procedures
    for rom in ("u8", "i8", "i16", "i32"):
        vals = rom_vals[rom]
        L.append("Private Sub silk_rom_fill_%s()" % rom)
        L.append("    Dim s As String, p() As String, i As Long")
        chunk, first = [], True
        for v in vals:
            chunk.append(str(v))
            if sum(len(c) + 1 for c in chunk) > 900:
                L.append('    s = %s"%s,"' % ("" if first else "s & ", ",".join(chunk)))
                first = False
                chunk = []
        if chunk:
            L.append('    s = %s"%s"' % ("" if first else "s & ", ",".join(chunk)))
        L.append("    p = Split(s, \",\")")
        L.append("    ReDim silk_rom_%s(0 To %d)" % (rom, len(vals) - 1))
        L.append("    For i = 0 To %d" % (len(vals) - 1))
        conv = {"u8": "CByte", "i8": "CInt", "i16": "CInt", "i32": "CLng"}[rom]
        L.append("        silk_rom_%s(i) = %s(p(i))" % (rom, conv))
        L.append("    Next")
        L.append("End Sub")
        L.append("")
    L.append("' (VB) fills the ROMs, the tables of pointers and the codebooks; the decoder's init calls it")
    L.append("Public Sub silk_tables_init()")
    L.append("    If silk_tables_ready Then Exit Sub")
    L.append("    silk_SigProc_FIX_init                   ' (VB) its tables of powers of 2")
    for rom in ("u8", "i8", "i16", "i32"):
        L.append("    silk_rom_fill_%s" % rom)
    for name, rom, targets in ptrs:
        for i, t in enumerate(targets):
            if offsets[t][0] != rom:
                sys.exit("%s: %s is not in rom %s" % (name, t, rom))
            L.append("    %s(%d) = %s" % (name, i, t))
    fieldnames = ["nVectors", "order", "quantStepSize_Q16", "invQuantStepSize_Q6", "CB1_NLSF_Q8", "CB1_Wght_Q9",
                  "CB1_iCDF", "pred_Q8", "ec_sel", "ec_iCDF", "ec_Rates_Q5", "deltaMin_Q15"]
    for i, (name, fields) in enumerate(structs):
        if len(fields) != len(fieldnames):
            sys.exit("%s: %d fields" % (name, len(fields)))
        L.append("    With silk_NLSF_CBs(%s)" % name)
        for fname, fv in zip(fieldnames, fields):
            if fv.startswith("silk_"):
                L.append("        .%s = %s" % (fname, fv))
            else:
                L.append("        .%s = %d%s' %s" % (fname, c_eval(fv), " " * 8, fv))
        L.append("    End With")
    L.append("    silk_tables_ready = True")
    L.append("End Sub")
    text = "\r\n".join(L) + "\r\n"
    data = text.encode("ascii")
    with open(OUT, "wb") as f:
        f.write(data)
    print("%s: %d tables (%s), %d pointer tables, %d codebooks" % (
        os.path.relpath(OUT, ROOT), len(tables), ", ".join("%s %d" % (r, len(v)) for r, v in rom_vals.items()),
        len(ptrs), len(structs)))


if __name__ == "__main__":
    main()
