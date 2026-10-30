"""Writes tools/ui_tests/campaign_walk.txt: campaign 1 from the main menu to its end. Each base or off-world
level is won with the campaign scripts' chat cheat ("cheat on", then "let me win") and its result screen's
Continue Game; each "between" level (SUB_x_xS: the base before an off-world mission) launches its transporter
as a player does: the units board it (the script's embark command, as clicking each onto it) and its launch
button is pressed (SUB_1_1S's transporter only comes once "let me win" sets its script's cheat flag).
python -I tools\\make_campaign_walk.py [levels to walk]"""
import sys

LEVELS = ["CAM_1A", "CAM_1B", "SUB_1_1S", "SUB_1_1", "SUB_1_2S", "SUB_1_2", "SUB_1_3S", "SUB_1_3", "CAM_1C",
          "CAM_1CA", "SUB_1_4AS", "SUB_1_4A", "SUB_1_5S", "SUB_1_5", "CAM_1A-C", "SUB_1_7S", "SUB_1_7",
          "SUB_1_DS", "SUB_1_D", "CAM_1END"]
LAUNCH = (29, 30)       # the transporter's launch button, at the top left
n = int(sys.argv[1]) if len(sys.argv) > 1 else len(LEVELS) - 1
lines = [
    "# campaign 1 from the main menu to its end (tools/make_campaign_walk.py writes this file): base and",
    "# off-world levels won with the campaign scripts' chat cheat ('cheat on', then 'let me win') and Continue",
    "# Game; each 'between' level (SUB_x_xS) launches its transporter: the units board it (embark), then launch",
    "wait 90", "click 512 365", "wait 60", "click 512 365", "wait 90", "click 607 575", "wait 900",
    "key 13", "wait 10", "text cheat on", "key 13", "wait 30",
]
for i, lev in enumerate(LEVELS[:n]):
    tag = "%02d_%s" % (i, lev.replace("-", "_"))
    lines += ["shot " + tag, "log level %s" % lev]
    if lev.endswith("S"):
        if lev == "SUB_1_1S":
            lines += ["key 13", "wait 10", "text let me win", "key 13"]
        lines += ["wait 900", "embark", "wait 600", "shot %s_boarded" % tag, "click %d %d" % LAUNCH, "wait 1500"]
    else:
        lines += ["key 13", "wait 10", "text let me win", "key 13", "wait 300", "shot %s_result" % tag,
                  "click 511 556", "wait 1200"]
lines += ["shot %02d_end" % n, "log done", "wait 5", "quit"]
open(r"C:\Users\Owner\desktop\websites\warzone\tools\ui_tests\campaign_walk.txt", "w").write("\n".join(lines) + "\n")
print("%d levels" % n)
