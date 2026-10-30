Attribute VB_Name = "aud_c"
' src/aud.cpp: what lib/sound (WzSound) asks the game (lib/sound/aud.h): an object's or a place's position
' in the sound's axes, whether an object is dead, the listener's position and rotation, and tracklib.h's
' sound_GetGameTime. SoundGame.cls (WzSound's ISoundGame) calls them. A SIMPLE_OBJECT * is the object's
' handle (0 = nullptr); a returned Vector3f is its fields ByRef. Also (VB) the exe's side of
' audio_PlayStream's std::function onFinished (audio_StreamFinished_new).
Option Explicit

Private Const M_PI_ As Double = 3.14159265358979    ' M_PI

Public Function audio_ObjectDead(ByVal psSimpleObj As Long) As Boolean
    ' check is valid simple object pointer
    If psSimpleObj = 0 Then
        debug_ 39, LOG_NEVER, "audio_ObjectDead", "audio_ObjectDead: simple object pointer invalid"
        audio_ObjectDead = True
        Exit Function
    End If

    ' check projectiles
    If isProjectile(psSimpleObj) Then
        audio_ObjectDead = aProj(castProjectile(psSimpleObj)).state = PROJ_POSTIMPACT
    Else
        ' check base object
        audio_ObjectDead = SIMPLE_OBJECT_died(psSimpleObj) <> 0
    End If
End Function

' @FIXME we don't need to do this, since we are not using qsound.
' Vector3f audio_GetPlayerPos(): the vector's fields
Public Sub audio_GetPlayerPos(x As Single, y As Single, z As Single)
    x = playerPos.p.x
    y = playerPos.p.z
    z = playerPos.p.y

    ' Invert Y to match QSOUND axes
    ' @NOTE What is QSOUND? Why invert the Y axis?
    y = world_coord(mapHeight) - y
End Sub

Public Sub audio_Get3DPlayerRotAboutVerticalAxis(angle As Single)
    angle = CSng(CDbl(CSng(playerPos.r.y) / DEG_1) * M_PI_ / 180#)
End Sub

Public Sub audio_GetStaticPos(ByVal iWorldX As Long, ByVal iWorldY As Long, piX As Long, piY As Long, piZ As Long)
    piX = iWorldX
    piZ = map_TileHeight(map_coord(iWorldX), map_coord(iWorldY))
    ' invert y to match QSOUND axes
    piY = world_coord(mapHeight) - iWorldY
End Sub

' @FIXME we don't need to do this, since we are not using qsound.
Public Sub audio_GetObjectPos(ByVal psBaseObj As Long, piX As Long, piY As Long, piZ As Long)
    Dim pos As Vector3i
    ' check is valid pointer
    If Not (psBaseObj <> 0) Then ASSERT_FAILURE "psBaseObj != nullptr", "src/aud.cpp:95", "audio_GetObjectPos", "Game object pointer invalid": Exit Sub

    pos = SIMPLE_OBJECT_pos(psBaseObj)
    piX = pos.x
    piZ = map_TileHeight(map_coord(pos.x), map_coord(pos.y))
    ' invert y to match QSOUND axes
    piY = world_coord(mapHeight) - pos.y
End Sub

Public Function sound_GetGameTime() As Long
    sound_GetGameTime = gameTime
End Function

' (VB) the std::function<void (const AUDIO_STREAM *, const void *)> of a lambda (a WzFunc, called with the
' stream and the user data) for audio_PlayStream: Nothing for an empty one
Public Function audio_StreamFinished_new(ByVal f As WzFunc) As IStreamFinished
    Dim h As SoundStreamFinished
    If f Is Nothing Then Exit Function
    Set h = New SoundStreamFinished
    Set h.f = f
    Set audio_StreamFinished_new = h
End Function
