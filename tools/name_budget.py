"""Estimate the VB6 name-table budget of a Warzone 2100 port, per area and per project.

VB6 keeps one case-insensitive name table per project, and every name counts
(procedures, parameters, locals, constants, UDT fields, module variables):
tools/probe/names_gen.py measured the ceiling at about 32,100 names. This
counts the distinct case-folded identifiers of the C++ that would be ported
(no comments, strings, C++ keywords or std::/glm::/fmt:: names; the files of
subsystems that are not ported are left out) and sums them for the planned
project split.

usage: python -I tools\\name_budget.py <warzone source dir>
"""
import os
import re
import sys

KEYWORDS = set("""alignas alignof and asm auto bool break case catch char char16_t char32_t class const constexpr
const_cast continue decltype default delete do double dynamic_cast else enum explicit export extern false float for
friend goto if inline int long mutable namespace new noexcept not nullptr operator or private protected public register
reinterpret_cast return short signed sizeof static static_assert static_cast struct switch template this thread_local
throw true try typedef typeid typename union unsigned using virtual void volatile wchar_t while override final define
ifdef ifndef endif elif undef pragma include defined size_t uint8_t uint16_t uint32_t uint64_t int8_t int16_t int32_t
int64_t intptr_t uintptr_t ptrdiff_t nullptr_t""".split())
LIBRARY_NAME = re.compile(r'\b(?:std|glm|nlohmann|fmt|tl|nonstd|optional_lite|gsl|chrono|filesystem)\s*::\s*([A-Za-z_]\w*)')
LIBRARY_PREFIXES = ('gl_', 'glad_', 'sdl_', 'al_', 'physfs')
IDENTIFIER = re.compile(r'\b[A-Za-z_][A-Za-z0-9_]*\b')

# Not ported (see docs/PLAN.md, "What is not ported")
NOT_PORTED = ['gfx_api_vk', 'gfx_api_null', 'quickjs_backend', 'wzscriptdebug', 'stdinreader', 'urlrequest',
              'emscripten', 'jpeg_encoder', 'png_util_spng', 'basis', 'exceptionhandler', 'integrations',
              'buildprovenance', 'lib/sequence', 'port_mapping', 'gns', 'GameNetworkingSockets', 'crashhandling']

AREAS = [('src/', 'src game+display'), ('lib/framework', 'framework'), ('lib/gamelib', 'gamelib'),
         ('lib/widget', 'widget'), ('lib/ivis_opengl', 'ivis'), ('lib/sdl', 'sdl'), ('lib/sound', 'sound'),
         ('lib/netplay', 'netplay'), ('lib/wzmaplib', 'wzmaplib')]
UI_DIRS = ('src/titleui/', 'src/hci/', 'src/screens/', 'src/researchtree/')
UI_FILES = ('hci.', 'intdisplay.', 'intelmap.', 'intorder.', 'intfac.', 'design.', 'frontend.', 'multiint.',
            'multimenu.', 'multilimit.', 'loadsave.', 'keyedit.', 'ingameop.', 'notifications.', 'wrappers.',
            'seqdisp.', 'chat.', 'challenge.', 'campaigninfo.', 'modding.', 'shortcuts.', 'difficulty.')
# The layout chosen in docs/PLAN.md: Warzone's own lib/ folders where they stand alone, the rest in the exe.
# A DLL's API names that the exe uses already appear in the exe's areas (src calls them).
PROJECTS = [('Warzone.exe', ['src game+display', 'src UI', 'gamelib', 'widget']),
            ('WzFramework.dll', ['framework']),
            ('WzGfx.dll', ['ivis']),
            ('WzSound.dll', ['sound']),
            ('WzMap.dll', ['wzmaplib']),
            ('WzNet.dll', ['netplay']),
            ('WzPlatform.dll', ['sdl']),
            ('(for comparison) all in one', ['src game+display', 'src UI', 'framework', 'gamelib', 'widget', 'ivis',
                                             'sdl', 'sound', 'netplay', 'wzmaplib'])]
CEILING = 32100


def strip(text):
    text = re.sub(r'/\*.*?\*/', ' ', text, flags=re.S)
    text = re.sub(r'//[^\n]*', ' ', text)
    text = re.sub(r'"(?:[^"\\\n]|\\.)*"', '""', text)
    return re.sub(r"'(?:[^'\\\n]|\\.)*'", "''", text)


def area_of(path, fname):
    area = next(a for prefix, a in AREAS if path.startswith(prefix))
    if area == 'src game+display' and (path.startswith(UI_DIRS) or fname.startswith(UI_FILES)):
        area = 'src UI'
    return area


def main(root):
    names = {}
    for top in ('src', 'lib'):
        for folder, _, files in os.walk(os.path.join(root, top)):
            rel = os.path.relpath(folder, root).replace(os.sep, '/')
            if '3rdparty' in rel:
                continue
            for fname in files:
                path = rel + '/' + fname
                if not fname.endswith(('.cpp', '.c', '.h', '.hpp')) or any(x in path for x in NOT_PORTED):
                    continue
                with open(os.path.join(folder, fname), encoding='utf-8', errors='replace') as f:
                    text = strip(f.read())
                library = set(n.lower() for n in LIBRARY_NAME.findall(text))
                bucket = names.setdefault(area_of(path, fname), set())
                for ident in IDENTIFIER.findall(text):
                    low = ident.lower()
                    if low not in KEYWORDS and low not in library and not low.startswith(LIBRARY_PREFIXES):
                        bucket.add(low)
    print("Distinct names per area (case-folded):")
    for area in sorted(names):
        print("  %-20s %6d" % (area, len(names[area])))
    print("  %-20s %6d" % ("all", len(set().union(*names.values()))))
    print("Per planned VB6 project (ceiling about %d):" % CEILING)
    for project, areas in PROJECTS:
        total = len(set().union(*(names.get(a, set()) for a in areas)))
        print("  %-20s %6d  (%d left for VB-side names)" % (project, total, CEILING - total))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
