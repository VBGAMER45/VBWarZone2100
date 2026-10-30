Attribute VB_Name = "movedef_h"
' src/movedef.h: the definitions for movement tracking.
' std::vector<Vector2i> asPath is the array asPath plus asPath_size. The
' C++ default member initialisers are all zero, as a fresh UDT is.
Option Explicit

Public Enum MOVE_STATUS
    MOVEINACTIVE
    MOVENAVIGATE
    MOVETURN
    MOVEPAUSE
    MOVEPOINTTOPOINT
    MOVETURNTOTARGET
    MOVEHOVER
    MOVEWAITROUTE
    MOVESHUFFLE
End Enum

Public Type MOVE_CONTROL
    Status As MOVE_STATUS               ' Inactive, Navigating or moving point to point status
    pathIndex As Long                   ' Position in asPath
    asPath() As Vector2i                ' Pointer to list of block X,Y map coordinates.
    asPath_size As Long
    destination As Vector2i             ' World coordinates of movement destination
    src As Vector2i
    target As Vector2i
    speed As Long                       ' Speed of motion
    moveDir As Long                     ' uint16_t: Direction of motion (not the direction the droid is facing)
    bumpDir As Long                     ' uint16_t: Direction at last bump
    bumpTime As Long                    ' unsigned: Time of first bump with something
    lastBump As Long                    ' uint16_t: Time of last bump with a droid - relative to bumpTime
    pauseTime As Long                   ' uint16_t: When MOVEPAUSE started - relative to bumpTime
    bumpPos As Vector3i                 ' Position of last bump
    tolerance As Long                   ' unsigned: Increases until unit gives up and goes to the next waypoint, if close enough
    shuffleStart As Long                ' unsigned: When a shuffle started
    psFormation As Long                 ' FORMATION *: formation the droid is currently a member of
    iVertSpeed As Long                  ' VTOL movement
End Type
