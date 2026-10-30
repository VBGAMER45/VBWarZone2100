Attribute VB_Name = "stubs_map"
' Stand-ins for the parts of wzmaplib not ported yet, with their C++ names
' and signatures. runMapScript (lib/wzmaplib/src/map_script.cpp, QuickJS)
' fails as if the script had: a map with a game.js does not load (as in
' tools/mapref.cpp, the reference build).
Option Explicit

' std::shared_ptr<Map> runMapScript(const std::vector<char>& fileBuffer, const std::string &path, uint32_t seed, bool preview, LoggingProtocol* logger)
Public Function runMapScript(fileBuffer() As Byte, ByVal fileSize As Long, ByVal path As String, ByVal seed As Long, ByVal preview As Boolean, ByVal logger As LoggingProtocol) As Map
    Set runMapScript = Nothing
End Function
