"""private_build.py - builds projects from src/ plus a staged folder into build\\j_<fork>\\bin, without touching src/,
bin/, compat/ or build/'s shared files: parallel forks build their work without seeing each other's (cluster J,
docs/PLAN.md section 37).

  python -I tools\\private_build.py <fork> <project>[,<project>...] --stage DIR [--with DIR:users ...] [--run ARGS...]

* project: WzFramework, WzMap, WzSound, WzBasis, WzGfx, WzNet, WzPlatform, WzSequence, Warzone (built in that order when
  several are given); a project not built here is referenced from build\\j_<fork>\\bin if the fork built it
  earlier, else from bin\\.
* stage (required): a mirror of src\\ holding the fork's new and changed files
  (<stage>\\net\\x.bas, <stage>\\game\\y.cls, <stage>\\shared\\...). A staged file named like a src\\ file replaces
  it; <stage>\\<folder>\\_deleted.txt lists src\\<folder> files (one name a line) left out of the build.
  src\\shared's folders (gfx, ogg, miniz, sodium) are compiled where build.py compiles them, a staged
  <stage>\\shared\\<name>\\x.bas replacing src's; new shared folders go in <stage>\\shared\\<name>\\ and need --with.
* --with DIR:users: a further folder of modules compiled into the listed project folders (comma-separated
  src folder names, e.g. --with C:\\...\\j_JZ\\src\\shared\\miniz:framework,net); a stage's own
  shared\\<name>\\ folders are passed the same way (a --with of a folder named like one of src\\shared's is
  ignored: that folder is compiled already, with the stage's files in it). DIR may be relative.
* No binary compatibility (private type libraries get fresh GUIDs); each private DLL is unregistered after its
  build. The exe gets its reg-free manifest from the private DLLs (build\\j_<fork>\\Warzone.res) and builds
  through a subst drive like build.py's (its LINK.EXE command line).
* --run ARGS: then runs build\\j_<fork>\\bin\\Warzone.exe ARGS (with --datadir=<tree>\\data added).
Copies bin\\'s DLLs the fork did not build into its bin\\ so the private exe runs."""
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402
import typelib  # noqa: E402

ORDER = [p[0] for p in build.PROJECTS]
REAL_BIN = build.BIN


def files_of(folder, stage_folder):
    """The project's source files: src's minus _deleted.txt and staged replacements, plus the staged ones."""
    deleted = set()
    dl = os.path.join(stage_folder, "_deleted.txt")
    if os.path.exists(dl):
        deleted = {ln.strip().lower() for ln in open(dl) if ln.strip()}
    staged = {}
    if os.path.isdir(stage_folder):
        for n in build.sources(stage_folder):
            staged[n.lower()] = os.path.join(stage_folder, n)
    out = []
    for n in build.sources(folder):
        if n.lower() in deleted or n.lower() in staged:
            continue
        out.append(os.path.join(folder, n))
    out += list(staged.values())
    return out


def main(argv):
    fork = argv[0]
    projects = argv[1].split(",")
    stage = None
    withs = []
    run = None
    i = 2
    while i < len(argv):
        if argv[i] == "--stage":
            stage = argv[i + 1]
            i += 2
        elif argv[i] == "--with":
            d, users = argv[i + 1].rsplit(":", 1)
            d = os.path.abspath(d)
            if os.path.basename(d).lower() in ("gfx", "ogg", "miniz", "sodium"):
                print("note: --with %s ignored: src\\shared\\%s is compiled already (with the stage's files)"
                      % (d, os.path.basename(d)))
            else:
                withs.append((d, users.split(",")))
            i += 2
        elif argv[i] == "--run":
            run = argv[i + 1:]
            break
        else:
            raise SystemExit("unknown argument " + argv[i])
    if stage is None:
        raise SystemExit("--stage DIR is required (a mirror of src\\ with the new and changed files)")
    staged_shared = os.path.join(stage, "shared")
    if os.path.isdir(staged_shared):
        for n in os.listdir(staged_shared):
            if os.path.isdir(os.path.join(staged_shared, n)) and n not in ("gfx", "ogg", "miniz", "sodium"):
                if not any(os.path.normcase(d) == os.path.normcase(os.path.join(staged_shared, n)) for d, _ in withs):
                    print("note: %s is not compiled anywhere without --with" % os.path.join(staged_shared, n))

    work = os.path.join(build.ROOT, "build", "j_" + fork)
    fbin = os.path.join(work, "bin")
    for d in (work, fbin):
        os.makedirs(d, exist_ok=True)
    # build.py's functions, pointed at the fork's folders
    build.BIN = fbin
    build.BUILD = work

    def dll_for(name):
        p = os.path.join(fbin, name + ".dll")
        return p if os.path.exists(p) and name in built_here else os.path.join(REAL_BIN, name + ".dll")

    built_here = set(n[:-4] for n in os.listdir(fbin) if n.lower().endswith(".dll") and
                     os.path.exists(os.path.join(work, n[:-4] + "_private.txt")))
    for name, kind, sub, refs in build.PROJECTS:
        if name not in projects:
            continue
        folder = os.path.join(build.SRC, sub)
        files = files_of(folder, os.path.join(stage, sub))
        shared_dirs = [(build.SHARED, None), (build.SHARED_GFX, build.SHARED_GFX_USERS),
                       (build.SHARED_OGG, build.SHARED_OGG_USERS), (build.SHARED_MINIZ, build.SHARED_MINIZ_USERS),
                       (build.SHARED_SODIUM, build.SHARED_SODIUM_USERS)]
        for sd, users in shared_dirs:
            if users is not None and sub not in users:
                continue
            files = [f for f in files] + files_of(sd, os.path.join(stage, os.path.relpath(sd, build.SRC)))
        for d, users in withs:
            if sub in users:
                files += [os.path.join(d, n) for n in build.sources(d)]
        lines = ["Type=" + ("OleDll" if kind == "dll" else "Exe"), build.STDOLE]
        for ref in refs:
            path = dll_for(ref)
            libname, tlbid, (major, minor), _ = typelib.library(path)
            lines.append("Reference=*\\G%s#%x.%x#0#%s#%s" % (tlbid, major, minor, path, libname))
        for p in files:
            build.normalise(p)
            if p.lower().endswith(".frm"):
                lines.append("Form=" + p)
            elif p.lower().endswith(".cls"):
                lines.append("Class=%s; %s" % (build.vb_name(p), p))
            else:
                lines.append("Module=%s; %s" % (build.vb_name(p), p))
        if kind == "exe":
            dlls = [r for r in refs]
            for r in dlls:                       # the manifest names the DLLs the exe runs with
                src_dll = dll_for(r)
                if os.path.normcase(os.path.dirname(src_dll)) != os.path.normcase(fbin):
                    shutil.copy(src_dll, fbin)
            res = build.write_manifest(dlls)
            if res is None:
                return 1
            lines.append('ResFile32="%s"' % res)
        lines.append(build.VBP_TAIL.format(name=name, ext="dll" if kind == "dll" else "exe", path=fbin,
                                           startup="(None)" if kind == "dll" else "Sub Main",
                                           compat='CompatibleMode="0"\n',
                                           bounds="-1" if name in build.OPTIMISED else "0",
                                           optimisation="0" if name in build.OPTIMISED else "2",
                                           threading="ThreadingModel=1\n" if kind == "dll" else ""))
        vbp_dir = os.path.join(work, "vbp")
        os.makedirs(vbp_dir, exist_ok=True)
        vbp = os.path.join(vbp_dir, name + ".vbp")
        with open(vbp, "w", newline="\r\n") as f:
            f.write("\n".join(lines))
        ok = build.make(name, vbp_dir, vbp)
        if not ok:
            return 1
        if kind == "dll":
            subprocess.run(["regsvr32", "/u", "/s", os.path.join(fbin, name + ".dll")])
            open(os.path.join(work, name + "_private.txt"), "w").write("built by j_build.py for " + fork + "\n")
            built_here.add(name)
    for n in os.listdir(REAL_BIN):                 # the rest of bin's DLLs, so the private exe runs
        if n.lower().endswith(".dll") and n[:-4] not in built_here and not os.path.exists(os.path.join(fbin, n)):
            shutil.copy(os.path.join(REAL_BIN, n), fbin)
    if run is not None:
        for a in run:                              # the game exits at once when --configdir's folder is missing
            if a.startswith("--configdir="):
                os.makedirs(a.split("=", 1)[1], exist_ok=True)
        exe = os.path.join(fbin, "Warzone.exe")
        r = subprocess.run([exe, "--datadir=" + os.path.join(build.ROOT, "data")] + run, cwd=fbin)
        print("Warzone.exe exit code %d" % r.returncode)
        return r.returncode
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
