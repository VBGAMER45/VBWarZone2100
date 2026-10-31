"""Writes a campaign walk UI test (tools/ui_tests/campaign_walk.txt for campaign 1, campaign2_walk.txt and
campaign3_walk.txt): the campaign from the main menu to its end. Each base or off-world level is won with the
campaign scripts' chat cheat ("cheat on", then "let me win") and its result screen's Continue Game; each
"between" level (SUB_x_xS, CAM_3_4S: the base before an off-world mission) launches its transporter as a
player does: the units board it (the script's embark command, as clicking each onto it) and its launch button
is pressed (SUB_1_1S's transporter only comes once "let me win" sets its script's cheat flag; campaigns 2 and
3 set theirs up at the start). Campaigns 2 and 3 start from the campaign selector's "Start at" drop-down.
The levels are libcampaign.js's __cam_alphaLevels, __cam_betaLevels, __cam_gammaLevels.
python -I tools\\make_campaign_walk.py [campaign 1-3] [levels to walk]
(campaign 1 walks to CAM_1END and stops there; campaigns 2 and 3 win their last level too: CAM_2END's win
loads CAM_3A, CAM_3_4's ends the game)"""
import sys

CAMPAIGNS = {
    1: ["CAM_1A", "CAM_1B", "SUB_1_1S", "SUB_1_1", "SUB_1_2S", "SUB_1_2", "SUB_1_3S", "SUB_1_3", "CAM_1C",
        "CAM_1CA", "SUB_1_4AS", "SUB_1_4A", "SUB_1_5S", "SUB_1_5", "CAM_1A-C", "SUB_1_7S", "SUB_1_7",
        "SUB_1_DS", "SUB_1_D", "CAM_1END"],
    2: ["CAM_2A", "SUB_2_1S", "SUB_2_1", "CAM_2B", "SUB_2_2S", "SUB_2_2", "CAM_2C", "SUB_2_5S", "SUB_2_5",
        "SUB_2DS", "SUB_2D", "SUB_2_6S", "SUB_2_6", "SUB_2_7S", "SUB_2_7", "SUB_2_8S", "SUB_2_8", "CAM_2END"],
    3: ["CAM_3A", "SUB_3_1S", "SUB_3_1", "CAM_3B", "SUB_3_2S", "SUB_3_2", "CAM3A-B", "CAM3C", "CAM3A-D1",
        "CAM3A-D2", "CAM_3_4S", "CAM_3_4"],
}
START_AT = {2: (562, 469), 3: (572, 497)}    # the selector's "Start at" drop-down items (Beta, Gamma Campaign)
LAUNCH = (29, 30)       # the transporter's launch button, at the top left
# frames to wait for the units to board (CAM_3_4S's landing zone, tile (50,245), is at the far end of
# CAM3A-D2's base map: the units need about a minute to drive there; 10 s left the transporter empty)
BOARD_WAIT = {"CAM_3_4S": 4200}
# frames to play before the win (campaign 3 started from the menu flies its units into CAM_3A: won before they
# land, SUB_3_1S starts without a truck or a factory and libcampaign's __camPlayerDead loses it at once)
PLAY_WAIT = {"CAM_3A": 2400}
camp = int(sys.argv[1]) if len(sys.argv) > 1 else 1
LEVELS = CAMPAIGNS[camp]
n = int(sys.argv[2]) if len(sys.argv) > 2 else (len(LEVELS) - 1 if camp == 1 else len(LEVELS))
lines = [
    "# campaign %d from the main menu to its end (tools/make_campaign_walk.py writes this file): base and" % camp,
    "# off-world levels won with the campaign scripts' chat cheat ('cheat on', then 'let me win') and Continue",
    "# Game; each 'between' level (SUB_x_xS) launches its transporter: the units board it (embark), then launch",
    "wait 90", "click 512 365", "wait 60", "click 512 365", "wait 90",
]
if camp in START_AT:
    lines += ["click 565 441", "wait 60", "click %d %d" % START_AT[camp], "wait 60", "shot selector"]
# Each level: waitlevel <name> (the level loaded and running, no video), then quiet 360 (6 s without a full-screen
# video: its briefing videos start a few seconds in, one after another); the cheat is typed once the first runs
lines += ["click 607 575"]
for i, lev in enumerate(LEVELS[:n]):
    tag = "%02d_%s" % (i, lev.replace("-", "_"))
    lines += ["waitlevel " + lev, "quiet 360"]
    if i == 0:
        lines += ["key 13", "wait 10", "text cheat on", "key 13", "wait 30"]
    if lev in PLAY_WAIT:
        lines += ["wait %d" % PLAY_WAIT[lev], "quiet 120"]
    lines += ["shot " + tag, "log level %s" % lev]
    if lev.endswith("S"):
        if lev == "SUB_1_1S":
            lines += ["key 13", "wait 10", "text let me win", "key 13"]
        # (SUB_1_1S's briefing video plays when the cheat's transporter arrives: quiet again before boarding, and
        # before the launch, which a video would swallow)
        lines += ["wait 300", "quiet 360", "embark", "wait %d" % BOARD_WAIT.get(lev, 600), "quiet 120",
                  "shot %s_boarded" % tag, "click %d %d" % LAUNCH, "wait 1200"]
    else:
        # (the win can start a video before the result screen: CAM_2C's cheat destroys the air base, which plays
        # MB2_C_MSG2; quiet again before Continue Game, which a video would swallow)
        lines += ["key 13", "wait 10", "text let me win", "key 13", "wait 300", "quiet 180",
                  "shot %s_result" % tag, "click 511 556", "wait 900"]
lines += ["quiet 360", "shot %02d_end" % n, "log done", "wait 5", "quit"]
name = "campaign_walk.txt" if camp == 1 else "campaign%d_walk.txt" % camp
open(r"C:\Users\Owner\desktop\websites\warzone\tools\ui_tests" + "\\" + name, "w").write("\n".join(lines) + "\n")
print("%s: %d levels" % (name, n))
