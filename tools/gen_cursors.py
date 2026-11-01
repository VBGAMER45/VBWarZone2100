"""
gen_cursors.py - writes src/platform/cursors_win_data.bas, the monochrome cursor images of lib/sdl/cursors_sdl.cpp
(its static const char *cursor_<name>[] arrays: the XPM header, three colours, 32 rows of 32 pixels and the hot
spot) and its cursors[CURSOR_MAX] table, from ref/warzone2100.

  python -I tools\\gen_cursors.py [out file]

Each image is a function returning its 37 lines joined by vbLf (VB has no initialised static arrays);
cursors_image(cur) is the table (cursors[cur].image).
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "ref", "warzone2100", "lib", "sdl", "cursors_sdl.cpp")
OUT = os.path.join(ROOT, "src", "platform", "cursors_win_data.bas")


def main(argv):
    out = argv[0] if argv else OUT
    text = open(SRC, encoding="utf-8").read()
    images = []
    for m in re.finditer(r"static const char \*(cursor_\w+)\[\] =\s*\{(.*?)\};", text, re.S):
        body = re.sub(r"/\*.*?\*/", "", m.group(2), flags=re.S)
        lines = re.findall(r'"([^"]*)"', body)
        assert len(lines) == 37, (m.group(1), len(lines))
        images.append((m.group(1), lines))
    table = re.search(r"cursors\[CURSOR_MAX\] =\s*\{(.*?)\};", text, re.S).group(1)
    entries = re.findall(r"\{\s*(cursor_\w+),\s*(CURSOR_\w+)\s*\}", table)
    vb = ['Attribute VB_Name = "cursors_win_data"',
          "' lib/sdl/cursors_sdl.cpp's cursor images (static const char *cursor_<name>[]: the XPM header, three",
          "' colours, 32 rows of 32 pixels, the hot spot) and its cursors[CURSOR_MAX] table. Written by",
          "' tools/gen_cursors.py from ref/warzone2100: do not edit.",
          "Option Explicit", ""]
    for name, lines in images:
        vb.append("Private Function %s() As String" % name)
        vb.append("    Dim s As String")
        for i, ln in enumerate(lines):
            vb.append('    s = %s"%s" & vbLf' % ("" if i == 0 else "s & ", ln))
        vb.append("    %s = s" % name)
        vb.append("End Function")
        vb.append("")
    vb.append("' cursors[cur].image: its lines (CURSOR order: cursors_win's constants)")
    vb.append("Public Function cursors_image(ByVal cur As Long) As String")
    vb.append("    Select Case cur")
    for i, (img, cur) in enumerate(entries):
        vb.append("    Case %d: cursors_image = %s()      ' %s" % (i, img, cur))
    vb.append("    End Select")
    vb.append("End Function")
    with open(out, "w", newline="\r\n", encoding="latin-1") as f:
        f.write("\n".join(vb) + "\n")
    print("%s: %d images, %d table entries" % (out, len(images), len(entries)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
