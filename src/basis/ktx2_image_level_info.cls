VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ktx2_image_level_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' basisu_transcoder.h: struct basist::ktx2_image_level_info, filled by ktx2_transcoder.get_image_level_info. A
' class (MultiUse: the caller makes one with New) since a UDT cannot cross the DLL's interface here; the uint32_t
' fields are Longs.
Option Explicit

' The mipmap level index (0=largest), texture array layer index, and cubemap face index of the image.
Public m_level_index As Long
Public m_layer_index As Long
Public m_face_index As Long

' The image's actual (or the original source image's) width/height in pixels, which may not be divisible by 4 pixels.
Public m_orig_width As Long
Public m_orig_height As Long

' The image's physical width/height, which will always be divisible by 4 pixels.
Public m_width As Long
Public m_height As Long

' The texture's dimensions in 4x4 texel blocks.
Public m_num_blocks_x As Long
Public m_num_blocks_y As Long

' The total number of blocks
Public m_total_blocks As Long

' true if the image has alpha data
Public m_alpha_flag As Boolean

' true if the image is an I-Frame. Currently, for ETC1S textures, the first frame will always be an I-Frame, and subsequent frames will always be P-Frames.
Public m_iframe_flag As Boolean
