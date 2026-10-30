// statsref_render.cpp - what loop.cpp's renderLoop and multiplay.cpp's multiPlayerLoop call in later clusters
// (the interface, the display, the sound, the network), as stubs_game.bas has them: each does nothing, or
// returns what stubs_game's stand-in returns, so that the real renderLoop and multiPlayerLoop run on both sides.
#include "lib/framework/frame.h"
#include "lib/framework/wzapp.h"
#include "lib/ivis_opengl/piedef.h"
#include "lib/ivis_opengl/piestate.h"
#include "lib/ivis_opengl/pieblitfunc.h"
#include "lib/ivis_opengl/screen.h"
#include "lib/sound/audio.h"
#include "lib/widget/widget.h"
#include "lib/netplay/netplay.h"
#include "src/display.h"
#include "src/hci.h"
#include "src/ingameop.h"
#include "src/loadsave.h"
#include "src/game.h"
#include "src/loop.h"
#include "src/mission.h"
#include "src/multiplay.h"
#include "src/multijoin.h"
#include "src/multistat.h"
#include "src/multivote.h"
#include "src/stdinreader.h"
#include "src/selection.h"
#include "src/console.h"
#include "src/edit3d.h"
#include "src/radar.h"
#include "src/notifications.h"
#include "src/keybind.h"
#include "src/wrappers.h"

// the globals (zero: nothing up, nothing dragged, the normal interface mode and mission state)
bool InGameOpUp = false;
bool isInGamePopupUp = false;
bool bLoadSaveUp = false;
bool bRequestLoad = false;
char sRequestResult[PATH_MAX] = {};
char saveGameName[256] = {};
INTMODE intMode = INT_NORMAL;
LOOP_MISSION_STATE loopMissionState = LMS_NORMAL;
LEVEL_TYPE nextMissionType = LEVEL_TYPE::LDS_NONE;
DragBox3D dragBox3D;
WallDrag wallDrag;

// gfx_api::context::get().shouldDraw(): stubs_game's gfx_api_context_shouldDraw (not headless: true)
bool statsref_shouldDraw()
{
	return !headlessGameMode();
}

// multiplay.cpp's static sendDataCheck2 (the network)
bool statsref_sendDataCheck2()
{
	return true;
}

// the interface (hci, ingameop, keybind, notifications, widgets)
void intAddInGamePopup() {}
void intDoScreenRefresh() {}
INT_RETVAL intRunWidgets() { return INT_NONE; }
void intDisplayWidgets() {}
bool intHelpOverlayIsUp() { return false; }
void intProcessInGameOptions(UDWORD) {}
void kf_TogglePauseMode() {}
bool isDraggingInGameNotification() { return false; }
bool isMouseClickDownOnScreenOverlayChild() { return false; }
bool getWidgetsStatus() { return true; }      // widget.cpp's bWidgetsActive = true
WidgetTriggers const &widgRunScreen(const std::shared_ptr<W_SCREEN> &psScreen)
{
	static WidgetTriggers none;
	return none;
}
void updateConsoleMessages() {}

// the display (display, display3d, radar, edit3d, lighting)
bool getRotActive() { return false; }
bool isMouseOverRadar() { return false; }
bool process3DBuilding() { return true; }
void processDeliveryRepos() {}
void displayRenderLoop() {}
void processInput() {}
void processMouseClickInput() {}
void displayWorld() {}

// the renderer, the sound and the platform
void screen_FlipIfBackDropTransition() {}
void pie_GetResetCounts(size_t *pPieCount, size_t *pPolyCount)
{
	*pPieCount = 0;
	*pPolyCount = 0;
}
void pie_SetFogStatus(bool val) {}
void pie_LoadBackDrop(SCREENTYPE screenType) {}
void wzPerfBegin(PERF_POINT pp, const char *descr) {}
void wzPerfEnd(PERF_POINT pp) {}
void audio_Update() {}
void audio_StopAll() {}
void wzSetCursor(CURSOR index) {}
void wzShowMouse(bool visible) {}

// the load/save screen, the save games and the missions
bool runLoadSave(bool bResetMissionWidgets) { return false; }
bool saveInMissionRes() { return false; }
bool saveMidMission() { return false; }
void deleteSaveGame_classic(const char *fileName) {}
const char *savegameWithoutExtension(const char *name) { return name; }
bool saveGame(const char *aFileName, GAME_TYPE saveType, bool isAutoSave) { return false; }
void missionDestroyObjects() {}
bool setUpMission(LEVEL_TYPE type) { return true; }
bool loop_GetVideoStatus() { return false; }
bool testPlayerHasWon() { return false; }

// the network: data checks, pings, the lobby's player info and identities, the kick routines, the win sequence
unsigned int selDroidDeselect(unsigned int player) { return 0; }
bool sendDataCheck() { return true; }
bool sendPing() { return true; }
void NETSendAllPlayerInfoTo(unsigned to) {}
bool sendMultiStatsHostVerifiedIdentities(uint32_t playerIndex) { return true; }
void wz_command_interface_output(const char *str, ...) {}
void wz_command_interface_output_room_status_json(bool queued) {}
void autoLagKickRoutine(std::chrono::steady_clock::time_point now) {}
void autoDesyncKickRoutine(std::chrono::steady_clock::time_point now) {}
void processPendingKickVotes() {}
bool multiplayerWinSequence(bool firstCall) { return true; }

// C2: what the messages (message.cpp), the gifts (multigifts.cpp, giftSingleDroid, giftSingleStructure) call in
// later clusters, as stubs_game.bas has them
#include "src/messagedef.h"
#include "src/message.h"
#include "src/qtscript.h"

// hci.cpp's proximity buttons without the widget (stubs_game's intAddProximityButton: the button's ID only, the
// asserts at hci.cpp's lines)
#define	IDPROX_START		120000		// The first proximity button
#define	IDPROX_END		129999		// The last proximity button
bool intAddProximityButton(PROXIMITY_DISPLAY *psProxDisp, UDWORD inc)
{
	UDWORD				cnt;

	if (selectedPlayer >= MAX_PLAYERS)
	{
		return false;
	}

	struct { UDWORD id; } sBFormInit;
	sBFormInit.id = IDPROX_START + inc;

	// loop back and find a free one!
	if (sBFormInit.id >= IDPROX_END)
	{
		for (cnt = IDPROX_START; cnt < IDPROX_END; cnt++)
		{
			// go down the prox msgs and see if it's free.
			const auto& proxDispList = apsProxDisp[selectedPlayer];
			auto proxDispIt = std::find_if(proxDispList.begin(), proxDispList.end(), [cnt](PROXIMITY_DISPLAY* pd)
			{
				return pd->buttonID == cnt;
			});

			if (proxDispIt == proxDispList.end())	// value was unused.
			{
				sBFormInit.id = cnt;
				break;
			}
		}
#line 2885 "src/hci.cpp"
		ASSERT_OR_RETURN(false, cnt != IDPROX_END, "Ran out of proximity displays");
	}
#line 2887 "src/hci.cpp"
	ASSERT_OR_RETURN(false, sBFormInit.id < IDPROX_END, "Invalid proximity message button ID %d", (int)sBFormInit.id);

	//store the ID so we can detect which one has been clicked on
	psProxDisp->buttonID = sBFormInit.id;

	return true;
}
#line 2912 "src/hci.cpp"
void intRemoveProximityButton(PROXIMITY_DISPLAY *psProxDisp)
{
	ASSERT_OR_RETURN(, psProxDisp->buttonID >= IDPROX_START && psProxDisp->buttonID <= IDPROX_END, "Invalid proximity ID");
}
#line 2000 "tools/statsref/statsref_render.cpp"

// qtscript.cpp's events (no script is loaded in the tests: the engine's triggers run over no instance)
bool triggerEventObjectTransfer(BASE_OBJECT *psObj, int from) { return true; }
bool triggerEventAllianceOffer(uint8_t from, uint8_t to) { return true; }
bool triggerEventAllianceAccepted(uint8_t from, uint8_t to) { return true; }
bool triggerEventAllianceBroken(uint8_t from, uint8_t to) { return true; }
