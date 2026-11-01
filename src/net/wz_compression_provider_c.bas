Attribute VB_Name = "wz_compression_provider_c"
' lib/netplay/wz_compression_provider.cpp, wz_compression_provider.h: class WzCompressionProvider, a singleton
' without state (Instance() is this module).
Option Explicit

' std::unique_ptr<ICompressionAdapter> newCompressionAdapter()
Public Function WzCompressionProvider_newCompressionAdapter() As ZlibCompressionAdapter
    ' Only support Zlib for the time being.
    ' TODO: in the future, we might want to support more compression algorithms, e.g. Zstd.
    ' In this case, we would need to somehow configure `WzCompressionAdapter` from
    ' the global WZ config to specify compression algorithm to use.
    Set WzCompressionProvider_newCompressionAdapter = New ZlibCompressionAdapter
End Function
