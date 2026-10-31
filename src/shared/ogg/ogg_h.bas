Attribute VB_Name = "ogg_h"
' libogg 1.3.6, include/ogg/ogg.h: the types (the functions: ogg_framing_c, framing.c). Compiled into the
' projects that read Ogg streams (tools/build.py's SHARED_OGG_USERS: WzSound for opusfile, WzSequence for the
' videos); module UDTs, so they never cross a DLL boundary (a public method takes their fields instead).
' C pointers become arrays the UDT owns: ogg_page's header and body and ogg_packet's packet are copies of the
' bytes the C points at (into the sync state's buffer and the stream state's body), made when the page or the
' packet is returned; ogg_sync_buffer returns the offset into oy.data where the caller writes. ogg_int64_t is a
' Double (exact to 2^53: granule positions and packet numbers of the game's streams stay far below).
Option Explicit

Public Type ogg_iovec_t
    iov_base() As Byte      ' void *iov_base
    iov_len As Long         ' size_t
End Type

' ogg_page is used to encapsulate the data in one Ogg bitstream page
Public Type ogg_page
    header() As Byte        ' unsigned char *header: (VB) a copy, from index 0
    header_len As Long
    body() As Byte          ' unsigned char *body: (VB) a copy, from index 0
    body_len As Long
End Type

' ogg_stream_state contains the current encode/decode state of a logical Ogg bitstream
Public Type ogg_stream_state
    body_data() As Byte     ' bytes from packet bodies
    body_storage As Long    ' storage elements allocated
    body_fill As Long       ' elements stored; fill mark
    body_returned As Long   ' elements of fill returned

    lacing_vals() As Long   ' The values that will go to the segment table
    granule_vals() As Double ' granulepos values for headers. Not compact this way, but it is simple coupled to
                            ' the lacing fifo
    lacing_storage As Long
    lacing_fill As Long
    lacing_packet As Long
    lacing_returned As Long

    header(0 To 281) As Byte ' working space for header encode
    header_fill As Long

    e_o_s As Long           ' set when we have buffered the last packet in the logical bitstream
    b_o_s As Long           ' set after we've written the initial page of a logical bitstream
    serialno As Long
    pageno As Long
    packetno As Double      ' sequence number for decode; the framing knows where there's a hole in the data,
                            ' but we need coupling so that the codec (which is in a separate abstraction
                            ' layer) also knows about the gap
    granulepos As Double
    vbInit As Boolean       ' (VB) the C tests body_data != NULL (ogg_stream_check): arrays allocated
End Type

' ogg_packet is used to encapsulate the data and metadata belonging to a single raw Ogg/Vorbis packet
Public Type ogg_packet
    packet() As Byte        ' unsigned char *packet: (VB) a copy, from index 0 (unallocated when bytes = 0)
    bytes As Long
    b_o_s As Long
    e_o_s As Long

    granulepos As Double

    packetno As Double      ' sequence number for decode; the framing knows where there's a hole in the data,
                            ' but we need coupling so that the codec (which is in a separate abstraction layer)
                            ' also knows about the gap
End Type

Public Type ogg_sync_state
    data() As Byte          ' unsigned char *data
    storage As Long
    fill As Long
    returned As Long

    unsynced As Long
    headerbytes As Long
    bodybytes As Long
    vbInit As Boolean       ' (VB) the C tests storage < 0 (ogg_sync_check) and data != NULL
End Type
