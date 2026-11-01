Attribute VB_Name = "cursors_win_data"
' lib/sdl/cursors_sdl.cpp's cursor images (static const char *cursor_<name>[]: the XPM header, three
' colours, 32 rows of 32 pixels, the hot spot) and its cursors[CURSOR_MAX] table. Written by
' tools/gen_cursors.py from ref/warzone2100: do not edit.
Option Explicit

Private Function cursor_arrow() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "X                               " & vbLf
    s = s & "XX                              " & vbLf
    s = s & "X.X                             " & vbLf
    s = s & "X..X                            " & vbLf
    s = s & "X...X                           " & vbLf
    s = s & "X....X                          " & vbLf
    s = s & "X.....X                         " & vbLf
    s = s & "X......X                        " & vbLf
    s = s & "X.......X                       " & vbLf
    s = s & "X........X                      " & vbLf
    s = s & "X.....XXXXX                     " & vbLf
    s = s & "X..X..X                         " & vbLf
    s = s & "X.X X..X                        " & vbLf
    s = s & "XX  X..X                        " & vbLf
    s = s & "X    X..X                       " & vbLf
    s = s & "     X..X                       " & vbLf
    s = s & "      X..X                      " & vbLf
    s = s & "      X..X                      " & vbLf
    s = s & "       XX                       " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "0,0" & vbLf
    cursor_arrow = s
End Function

Private Function cursor_dest() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "           ..                   " & vbLf
    s = s & "           ...X                 " & vbLf
    s = s & "           ..X.X                " & vbLf
    s = s & "           ..X .X               " & vbLf
    s = s & "           .X.X .X              " & vbLf
    s = s & "           .X.X  .X             " & vbLf
    s = s & "           .X .X  .X            " & vbLf
    s = s & "           .X .X   .......      " & vbLf
    s = s & "           .X  .X  .XX....X     " & vbLf
    s = s & "           .X  .X .XXX....X     " & vbLf
    s = s & "           .X   ..XXXX....X     " & vbLf
    s = s & "          ...X  ..XXXX....X     " & vbLf
    s = s & "         .....X ..........X     " & vbLf
    s = s & "         .....X ..........X     " & vbLf
    s = s & "         .....X .XX.XX.XX..     " & vbLf
    s = s & "          ...X .X..X..X..X..    " & vbLf
    s = s & "           XX  .X..X..X..XX.X   " & vbLf
    s = s & "               X...........XX   " & vbLf
    s = s & "                XXXXXXXXXXXX    " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,18" & vbLf
    cursor_dest = s
End Function

Private Function cursor_sight() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "            .......             " & vbLf
    s = s & "   ..     ...........     ..    " & vbLf
    s = s & "   ....  X..XXXXXXX..X  ....X   " & vbLf
    s = s & "    .....XXXX      XXX.....XX   " & vbLf
    s = s & "    .......X        .......X    " & vbLf
    s = s & "     ........     ........XX    " & vbLf
    s = s & "     .......XX     .......X     " & vbLf
    s = s & "    XX.....XX       .....XXX    " & vbLf
    s = s & "   ..X....XX         ....X..X   " & vbLf
    s = s & "   ..XX..XX           ..XX..X   " & vbLf
    s = s & "  ...X .XX             .X ...X  " & vbLf
    s = s & "  ..XX  X               X  ..X  " & vbLf
    s = s & "  ..X                      ..X  " & vbLf
    s = s & "  ..X                      ..X  " & vbLf
    s = s & "  ..X                      ..X  " & vbLf
    s = s & "  ..X                      ..X  " & vbLf
    s = s & "  ...  .               .X ...X  " & vbLf
    s = s & "   ..X ..             ..X ..XX  " & vbLf
    s = s & "   ..X....           ....X..X   " & vbLf
    s = s & "    XX.....         .....XXXX   " & vbLf
    s = s & "     .......       .......XX    " & vbLf
    s = s & "     ........     ........XX    " & vbLf
    s = s & "    .......XXX    XX.......X    " & vbLf
    s = s & "    .....XXX        XX.....X    " & vbLf
    s = s & "   ....XXX...     ...XXX....    " & vbLf
    s = s & "   ..XXX X...........XXX X..X   " & vbLf
    s = s & "    XX    XX.......XXX     XX   " & vbLf
    s = s & "            XXXXXXXX            " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,16" & vbLf
    cursor_sight = s
End Function

Private Function cursor_target() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "               .                " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "            ....X..             " & vbLf
    s = s & "          ......X....           " & vbLf
    s = s & "         ...XXX.XXX...          " & vbLf
    s = s & "        ..XX   .X   X..         " & vbLf
    s = s & "       ..X     .X     ..        " & vbLf
    s = s & "      ..X      XX      ..       " & vbLf
    s = s & "      ..X              ..X      " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "   ........X   .    ........X   " & vbLf
    s = s & "    XXXXXXXX   X     XXXXXXXX   " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "      ..                .X      " & vbLf
    s = s & "      ..X      .       ..X      " & vbLf
    s = s & "       ..      .X     ..X       " & vbLf
    s = s & "        ..     .X    ..X        " & vbLf
    s = s & "         ...   .X   ..X         " & vbLf
    s = s & "         X......X....X          " & vbLf
    s = s & "          XX....X..XX           " & vbLf
    s = s & "            XXX.XXX             " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "               XX               " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_target = s
End Function

Private Function cursor_larrow() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                ..              " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "            ....X               " & vbLf
    s = s & "          ...............       " & vbLf
    s = s & "        .................X      " & vbLf
    s = s & "      ...................X      " & vbLf
    s = s & "    .....................X      " & vbLf
    s = s & "    XX...................X      " & vbLf
    s = s & "      XX.................X      " & vbLf
    s = s & "        XX...............X      " & vbLf
    s = s & "          XX....XXXXXXXXXX      " & vbLf
    s = s & "            XX...X              " & vbLf
    s = s & "              XX..X             " & vbLf
    s = s & "                XXX             " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "7,15" & vbLf
    cursor_larrow = s
End Function

Private Function cursor_rarrow() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "              ..                " & vbLf
    s = s & "               ...              " & vbLf
    s = s & "                ....            " & vbLf
    s = s & "       ...............          " & vbLf
    s = s & "       .................        " & vbLf
    s = s & "       ...................      " & vbLf
    s = s & "       .....................X   " & vbLf
    s = s & "       ...................XX    " & vbLf
    s = s & "       .................XX      " & vbLf
    s = s & "       ...............XX        " & vbLf
    s = s & "        XXXXXXXX....XX          " & vbLf
    s = s & "              X...XX            " & vbLf
    s = s & "              ..XX              " & vbLf
    s = s & "              XX                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "25,16" & vbLf
    cursor_rarrow = s
End Function

Private Function cursor_darrow() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "            .......             " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X XX         " & vbLf
    s = s & "         .X .......XX.X         " & vbLf
    s = s & "         ..X.......X..X         " & vbLf
    s = s & "          ...........X          " & vbLf
    s = s & "          ...........X          " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,24" & vbLf
    cursor_darrow = s
End Function

Private Function cursor_uarrow() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "               .                " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "          ...........X          " & vbLf
    s = s & "          ...........X          " & vbLf
    s = s & "         ..X.......X..X         " & vbLf
    s = s & "         .X .......XX.X         " & vbLf
    s = s & "            .......X XX         " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "             XXXXXXX            " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,8" & vbLf
    cursor_uarrow = s
End Function

Private Function cursor_default() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "            ..XX                " & vbLf
    s = s & "            ....XX              " & vbLf
    s = s & "             .....XX            " & vbLf
    s = s & "             .......XX          " & vbLf
    s = s & "              ........XX        " & vbLf
    s = s & "              ..........XX      " & vbLf
    s = s & "               ...........X     " & vbLf
    s = s & "               ....XXXXXXXX     " & vbLf
    s = s & "                ...X            " & vbLf
    s = s & "                ...X            " & vbLf
    s = s & "                 ..X            " & vbLf
    s = s & "                 ..X            " & vbLf
    s = s & "                  .X            " & vbLf
    s = s & "                  .X            " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "12,12" & vbLf
    cursor_default = s
End Function

Private Function cursor_attach() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "     ...X  ....X  ....X  ...    " & vbLf
    s = s & "     XXX.X.XXXX.X.XXXX.X.XXX    " & vbLf
    s = s & "       .....  .....  .....      " & vbLf
    s = s & "        .X.X   .X.X   .X.X      " & vbLf
    s = s & "     ...X  ....X  ....X  ...    " & vbLf
    s = s & "     XXX   XXXX   XXXX   XXX    " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "17,14" & vbLf
    cursor_attach = s
End Function

Private Function cursor_attack() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "               .                " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "            ....X..             " & vbLf
    s = s & "          ......X....           " & vbLf
    s = s & "         ...XXX.XXX...          " & vbLf
    s = s & "        ..XX   .X   X..         " & vbLf
    s = s & "       ..X     .X     ..        " & vbLf
    s = s & "      ..X      XX      ..       " & vbLf
    s = s & "      ..X              ..X      " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "   ........X   .    ........X   " & vbLf
    s = s & "    XXXXXXXX   X     XXXXXXXX   " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "     ..X                ..X     " & vbLf
    s = s & "      ..                .X      " & vbLf
    s = s & "      ..X      .       ..X      " & vbLf
    s = s & "       ..      .X     ..X       " & vbLf
    s = s & "        ..     .X    ..X        " & vbLf
    s = s & "         ...   .X   ..X         " & vbLf
    s = s & "         X......X....X          " & vbLf
    s = s & "          XX....X..XX           " & vbLf
    s = s & "            XXX.XXX             " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "               XX               " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_attack = s
End Function

Private Function cursor_bomb() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                   . .          " & vbLf
    s = s & "                  . . .         " & vbLf
    s = s & "                  XXX.          " & vbLf
    s = s & "               XXXX . .         " & vbLf
    s = s & "               XX  . .          " & vbLf
    s = s & "               XX               " & vbLf
    s = s & "              ....X             " & vbLf
    s = s & "              ....X             " & vbLf
    s = s & "            .XXXXXX.XX          " & vbLf
    s = s & "          ............X         " & vbLf
    s = s & "         ..............X        " & vbLf
    s = s & "        ................X       " & vbLf
    s = s & "       ..................X      " & vbLf
    s = s & "       ..................X      " & vbLf
    s = s & "      ....................X     " & vbLf
    s = s & "      ....................X     " & vbLf
    s = s & "      ....................X     " & vbLf
    s = s & "      ....................X     " & vbLf
    s = s & "      ....................X     " & vbLf
    s = s & "       ..................X      " & vbLf
    s = s & "       ..................X      " & vbLf
    s = s & "        ................X       " & vbLf
    s = s & "         ..............X        " & vbLf
    s = s & "          ............X         " & vbLf
    s = s & "            ........XX          " & vbLf
    s = s & "            XXXXXXXX            " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "16,16" & vbLf
    cursor_bomb = s
End Function

Private Function cursor_bridge() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "           ..                   " & vbLf
    s = s & "         ....                   " & vbLf
    s = s & "       .......                  " & vbLf
    s = s & "     .........                  " & vbLf
    s = s & "   ............                 " & vbLf
    s = s & "   ............             ..  " & vbLf
    s = s & "   X............          .. .  " & vbLf
    s = s & "   X............        ..XX.   " & vbLf
    s = s & "    X............     ..XXX.    " & vbLf
    s = s & "    X............   ..XXXX.     " & vbLf
    s = s & "     X..............XXXXX.      " & vbLf
    s = s & "   ..X............XXXXXX.       " & vbLf
    s = s & "  ....X..........XXXXX..        " & vbLf
    s = s & "  ....X.........XXXX......      " & vbLf
    s = s & "  .....X.......XXX.........     " & vbLf
    s = s & "  .....X......XX............    " & vbLf
    s = s & "  ......X....XXXXX..........X   " & vbLf
    s = s & "  ......X...XX..XXXX........X   " & vbLf
    s = s & " ........X. ..XXX    .......X   " & vbLf
    s = s & ".........X..XXX      .......X.. " & vbLf
    s = s & " XX......X XX        ..........X" & vbLf
    s = s & "   XXX...X           .........XX" & vbLf
    s = s & "      XXXX           .........X " & vbLf
    s = s & "                     ........XX " & vbLf
    s = s & "                   ..........X  " & vbLf
    s = s & "                    XX......XX  " & vbLf
    s = s & "                      XXX...X   " & vbLf
    s = s & "                         XXXX   " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "16,16" & vbLf
    cursor_bridge = s
End Function

Private Function cursor_build() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                ......          " & vbLf
    s = s & "              .....XX..X        " & vbLf
    s = s & "           .X.....X  XXX        " & vbLf
    s = s & "          .X......X             " & vbLf
    s = s & "          X.......X             " & vbLf
    s = s & "          .......XX             " & vbLf
    s = s & "         .......X               " & vbLf
    s = s & "         ......X.X              " & vbLf
    s = s & "       .......X...X             " & vbLf
    s = s & "      .......X ....X            " & vbLf
    s = s & "       ....XX   ....X           " & vbLf
    s = s & "        ...X     ....X          " & vbLf
    s = s & "         .X       ....X         " & vbLf
    s = s & "                   ....X        " & vbLf
    s = s & "                    ....X       " & vbLf
    s = s & "                     ...X       " & vbLf
    s = s & "                      ..X       " & vbLf
    s = s & "                       X        " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_build = s
End Function

Private Function cursor_embark() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "        .................X      " & vbLf
    s = s & "         ...............XX      " & vbLf
    s = s & "          .............X        " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           X.........X          " & vbLf
    s = s & "        ...XX.......XX...X      " & vbLf
    s = s & "         ...XX.....XX...X       " & vbLf
    s = s & "          ...XX...XX...X        " & vbLf
    s = s & "           ...XX.XX...X         " & vbLf
    s = s & "            ...XXX...X          " & vbLf
    s = s & "             ...X...X           " & vbLf
    s = s & "              .....X            " & vbLf
    s = s & "               ...X             " & vbLf
    s = s & "                .X              " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "16,22" & vbLf
    cursor_embark = s
End Function

Private Function cursor_disembark() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "        .................X      " & vbLf
    s = s & "         ...............XX      " & vbLf
    s = s & "          .............X        " & vbLf
    s = s & "           ...........X         " & vbLf
    s = s & "           X.........X          " & vbLf
    s = s & "        ...XX.......XX...X      " & vbLf
    s = s & "         ...XX.....XX...X       " & vbLf
    s = s & "          ...XX...XX...X        " & vbLf
    s = s & "           ...XX.XX...X         " & vbLf
    s = s & "            ...XXX...X          " & vbLf
    s = s & "             ...X...X           " & vbLf
    s = s & "              .....X            " & vbLf
    s = s & "               ...X             " & vbLf
    s = s & "                .X              " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "16,22" & vbLf
    cursor_disembark = s
End Function

Private Function cursor_fix() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "           ...                  " & vbLf
    s = s & "             ..X                " & vbLf
    s = s & "              ..X               " & vbLf
    s = s & "              ..X               " & vbLf
    s = s & "       .      ..X               " & vbLf
    s = s & "       .X     ..X               " & vbLf
    s = s & "       .X    X..X               " & vbLf
    s = s & "       ..XXXX...X               " & vbLf
    s = s & "        .........X              " & vbLf
    s = s & "         .........X             " & vbLf
    s = s & "              .....X            " & vbLf
    s = s & "               .....XXXXX       " & vbLf
    s = s & "                .........X      " & vbLf
    s = s & "                 .........X     " & vbLf
    s = s & "                  ...X   ..X    " & vbLf
    s = s & "                  ..X     .X    " & vbLf
    s = s & "                  ..X     .X    " & vbLf
    s = s & "                  ..X           " & vbLf
    s = s & "                  ..X           " & vbLf
    s = s & "                   ..XXX        " & vbLf
    s = s & "                    ...X        " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "16,16" & vbLf
    cursor_fix = s
End Function

Private Function cursor_guard() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "       .....X      .....        " & vbLf
    s = s & "       .XXX.........XXX.X       " & vbLf
    s = s & "       .X..XXXXXXXXX..X.X       " & vbLf
    s = s & "       .X.............X.X       " & vbLf
    s = s & "       .X.............X.X       " & vbLf
    s = s & "       .X.............X.X       " & vbLf
    s = s & "       .X............X.X        " & vbLf
    s = s & "        .X...........X.X        " & vbLf
    s = s & "        .X...........X.X        " & vbLf
    s = s & "        .XX.........XX.X        " & vbLf
    s = s & "         .X.........X.X         " & vbLf
    s = s & "         .XX.......XX.X         " & vbLf
    s = s & "          .XX.....XX.X          " & vbLf
    s = s & "           .XX...XX.X           " & vbLf
    s = s & "            ..XXX..X            " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "               .XX              " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "16,17" & vbLf
    cursor_guard = s
End Function

Private Function cursor_jam() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "           ...XXX...X           " & vbLf
    s = s & "          ..XX   XX..X          " & vbLf
    s = s & "         ..X       X..X         " & vbLf
    s = s & "         .X   ...X  X.X         " & vbLf
    s = s & "        ..X  .X  .X  ..X        " & vbLf
    s = s & "        .X  .X    .X  .X        " & vbLf
    s = s & "        .X  .X .X .X  .X        " & vbLf
    s = s & "        .X  .XX.XX.X  .X        " & vbLf
    s = s & "        ..X  X...XX  ..X        " & vbLf
    s = s & "         .X   ...X   .X         " & vbLf
    s = s & "         ..X .....X ..X         " & vbLf
    s = s & "          .XX.....XX.X          " & vbLf
    s = s & "           X.......XX           " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "            XXXXXXXXX           " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_jam = s
End Function

Private Function cursor_lockon() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "           ...XXX...X           " & vbLf
    s = s & "          ..XX   XX..X          " & vbLf
    s = s & "         ..X       X..X         " & vbLf
    s = s & "         .X   ...X  X.X         " & vbLf
    s = s & "        ..X  .X  .X  ..X        " & vbLf
    s = s & "        .X  .X    .X  .X        " & vbLf
    s = s & "        .X  .X .X .X  .X        " & vbLf
    s = s & "        .X  .XX.XX.X  .X        " & vbLf
    s = s & "        ..X  X...XX  ..X        " & vbLf
    s = s & "         .X   ...X   .X         " & vbLf
    s = s & "         ..X .....X ..X         " & vbLf
    s = s & "          .XX.....XX.X          " & vbLf
    s = s & "           X.......XX           " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "            XXXXXXXXX           " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_lockon = s
End Function

Private Function cursor_scout() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "             .....X             " & vbLf
    s = s & "           ...XXX...X           " & vbLf
    s = s & "          ..XX   XX..X          " & vbLf
    s = s & "         ..X       X..X         " & vbLf
    s = s & "         .X   ...X  X.X         " & vbLf
    s = s & "        ..X  .X  .X  ..X        " & vbLf
    s = s & "        .X  .X    .X  .X        " & vbLf
    s = s & "        .X  .X .X .X  .X        " & vbLf
    s = s & "        .X  .XX.XX.X  .X        " & vbLf
    s = s & "        ..X  X...XX  ..X        " & vbLf
    s = s & "         .X   ...X   .X         " & vbLf
    s = s & "         ..X .....X ..X         " & vbLf
    s = s & "          .XX.....XX.X          " & vbLf
    s = s & "           X.......XX           " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "            XXXXXXXXX           " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_scout = s
End Function

Private Function cursor_menu() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "          ..XX                  " & vbLf
    s = s & "          ....XX                " & vbLf
    s = s & "           .....XX              " & vbLf
    s = s & "           .......XX            " & vbLf
    s = s & "            ........XX          " & vbLf
    s = s & "            ..........XX        " & vbLf
    s = s & "             ........XX         " & vbLf
    s = s & "             .....XX            " & vbLf
    s = s & "              ...XX             " & vbLf
    s = s & "              ...XX             " & vbLf
    s = s & "               ..XX             " & vbLf
    s = s & "               . X              " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "11,11" & vbLf
    cursor_menu = s
End Function

Private Function cursor_move() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "           .  .....X .          " & vbLf
    s = s & "          ..X.XXXXX.X..         " & vbLf
    s = s & "         ...XXX    XX...        " & vbLf
    s = s & "        ....X.X    .X....       " & vbLf
    s = s & "       .....X.X    .X.....      " & vbLf
    s = s & "        XXXXX.X    .XXXXX       " & vbLf
    s = s & "        .X....X    ....X.X      " & vbLf
    s = s & "       .XXXXXXX          .X     " & vbLf
    s = s & "       .X                .X     " & vbLf
    s = s & "       .X                .X     " & vbLf
    s = s & "       .X                .X     " & vbLf
    s = s & "       .X                .X     " & vbLf
    s = s & "        .X....     ....X.X      " & vbLf
    s = s & "        XXXXX.X    .XXXXX       " & vbLf
    s = s & "       .....X.X    .X.....      " & vbLf
    s = s & "        ....X.X    .X....X      " & vbLf
    s = s & "         ...XXX    XX...X       " & vbLf
    s = s & "          ..X.X    .X..X        " & vbLf
    s = s & "           .X .....X .X         " & vbLf
    s = s & "              XXXXX             " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_move = s
End Function

Private Function cursor_notpossible() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "        .X            .X        " & vbLf
    s = s & "       ...X          ...X       " & vbLf
    s = s & "      .....X        .....X      " & vbLf
    s = s & "     .......X      .......X     " & vbLf
    s = s & "      .......X    .......X      " & vbLf
    s = s & "       .......X  .......X       " & vbLf
    s = s & "        .......X.......X        " & vbLf
    s = s & "         .............X         " & vbLf
    s = s & "          ...........X          " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "            .......X            " & vbLf
    s = s & "           .........X           " & vbLf
    s = s & "          ...........X          " & vbLf
    s = s & "         .............X         " & vbLf
    s = s & "        .......X.......X        " & vbLf
    s = s & "       .......X  .......X       " & vbLf
    s = s & "      .......X    .......X      " & vbLf
    s = s & "     .......X      .......X     " & vbLf
    s = s & "      .....X        .....X      " & vbLf
    s = s & "       ...X          ...X       " & vbLf
    s = s & "        .X            .X        " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,17" & vbLf
    cursor_notpossible = s
End Function

Private Function cursor_pickup() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "             .X  .X             " & vbLf
    s = s & "             .X  .X             " & vbLf
    s = s & "             .X.X.X             " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "             .X.X.X             " & vbLf
    s = s & "             .X  .X             " & vbLf
    s = s & "             .X.X.X             " & vbLf
    s = s & "              ...X              " & vbLf
    s = s & "              X.XX              " & vbLf
    s = s & "             .X.X.X             " & vbLf
    s = s & "            ..X.X..X            " & vbLf
    s = s & "           ...X.X...X           " & vbLf
    s = s & "          ....X.X....X          " & vbLf
    s = s & "         .....XXX.....X         " & vbLf
    s = s & "         ...XXXXXXX...X         " & vbLf
    s = s & "         .XXX      XX.X         " & vbLf
    s = s & "         .X          .X         " & vbLf
    s = s & "         .X          .X         " & vbLf
    s = s & "         .X          .X         " & vbLf
    s = s & "         .X          .X         " & vbLf
    s = s & "         ..X        ..X         " & vbLf
    s = s & "          ..X      ..X          " & vbLf
    s = s & "           ..X    ..X           " & vbLf
    s = s & "            .X    .X            " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,20" & vbLf
    cursor_pickup = s
End Function

Private Function cursor_seekrepair() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "      ...                       " & vbLf
    s = s & "        ..X             X       " & vbLf
    s = s & "         ..X            .X      " & vbLf
    s = s & "         ..X            ..X     " & vbLf
    s = s & "  .      ..X        XXXX...X    " & vbLf
    s = s & "  .X     ..X        ........X   " & vbLf
    s = s & "  .X    X..X        .........X  " & vbLf
    s = s & "  ..XXXX...X        ..........  " & vbLf
    s = s & "   .........X       .........   " & vbLf
    s = s & "    .........X      ........    " & vbLf
    s = s & "         .....X         ...     " & vbLf
    s = s & "          .....XXXXX    ..      " & vbLf
    s = s & "           .........X   .       " & vbLf
    s = s & "            .........X          " & vbLf
    s = s & "             ...X   ..X         " & vbLf
    s = s & "             ..X     .X         " & vbLf
    s = s & "             ..X     .X         " & vbLf
    s = s & "             ..X                " & vbLf
    s = s & "             ..X                " & vbLf
    s = s & "              ..XXX             " & vbLf
    s = s & "               ...X             " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "16,15" & vbLf
    cursor_seekrepair = s
End Function

Private Function cursor_select() As String
    Dim s As String
    s = "    32    32        3            1" & vbLf
    s = s & "X c #000000" & vbLf
    s = s & ". c #ffffff" & vbLf
    s = s & "  c None" & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "    .....X            .....X    " & vbLf
    s = s & "    .XXXXX             XXX.X    " & vbLf
    s = s & "    .X                    .X    " & vbLf
    s = s & "    .X                    .X    " & vbLf
    s = s & "    .X                    .X    " & vbLf
    s = s & "    XX                    XX    " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "               .X               " & vbLf
    s = s & "               XX               " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "    .                     .     " & vbLf
    s = s & "    .X                    .X    " & vbLf
    s = s & "    .X                    .X    " & vbLf
    s = s & "    .X                    .X    " & vbLf
    s = s & "    .....X            .....X    " & vbLf
    s = s & "    XXXXXX            XXXXXX    " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "                                " & vbLf
    s = s & "15,15" & vbLf
    cursor_select = s
End Function

' cursors[cur].image: its lines (CURSOR order: cursors_win's constants)
Public Function cursors_image(ByVal cur As Long) As String
    Select Case cur
    Case 0: cursors_image = cursor_arrow()      ' CURSOR_ARROW
    Case 1: cursors_image = cursor_dest()      ' CURSOR_DEST
    Case 2: cursors_image = cursor_sight()      ' CURSOR_SIGHT
    Case 3: cursors_image = cursor_target()      ' CURSOR_TARGET
    Case 4: cursors_image = cursor_larrow()      ' CURSOR_LARROW
    Case 5: cursors_image = cursor_rarrow()      ' CURSOR_RARROW
    Case 6: cursors_image = cursor_darrow()      ' CURSOR_DARROW
    Case 7: cursors_image = cursor_uarrow()      ' CURSOR_UARROW
    Case 8: cursors_image = cursor_default()      ' CURSOR_DEFAULT
    Case 9: cursors_image = cursor_default()      ' CURSOR_EDGEOFMAP
    Case 10: cursors_image = cursor_attach()      ' CURSOR_ATTACH
    Case 11: cursors_image = cursor_attack()      ' CURSOR_ATTACK
    Case 12: cursors_image = cursor_bomb()      ' CURSOR_BOMB
    Case 13: cursors_image = cursor_bridge()      ' CURSOR_BRIDGE
    Case 14: cursors_image = cursor_build()      ' CURSOR_BUILD
    Case 15: cursors_image = cursor_embark()      ' CURSOR_EMBARK
    Case 16: cursors_image = cursor_disembark()      ' CURSOR_DISEMBARK
    Case 17: cursors_image = cursor_fix()      ' CURSOR_FIX
    Case 18: cursors_image = cursor_guard()      ' CURSOR_GUARD
    Case 19: cursors_image = cursor_jam()      ' CURSOR_JAM
    Case 20: cursors_image = cursor_lockon()      ' CURSOR_LOCKON
    Case 21: cursors_image = cursor_scout()      ' CURSOR_SCOUT
    Case 22: cursors_image = cursor_menu()      ' CURSOR_MENU
    Case 23: cursors_image = cursor_move()      ' CURSOR_MOVE
    Case 24: cursors_image = cursor_notpossible()      ' CURSOR_NOTPOSSIBLE
    Case 25: cursors_image = cursor_pickup()      ' CURSOR_PICKUP
    Case 26: cursors_image = cursor_seekrepair()      ' CURSOR_SEEKREPAIR
    Case 27: cursors_image = cursor_select()      ' CURSOR_SELECT
    End Select
End Function
