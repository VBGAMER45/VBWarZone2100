Attribute VB_Name = "celt_modes_c"
' ref/opus/celt (libopus 1.5.2): modes.c, modes.h, the mode struct of static_modes_float.h, kiss_fft.h's and
' mdct.h's structs (cluster K, fork O2). CUSTOM_MODES is undefined: opus_custom_mode_create only finds the
' static 48 kHz / 960 mode (mode48000_960_120 of static_modes_float.h, the float build's); compute_ebands,
' compute_allocation_table, opus_custom_mode_destroy and the CUSTOM_MODES part of opus_custom_mode_create are
' not compiled in the C either. The tables are celt_tables_vb's (generated from the C text).
' (VB) A const CELTMode * is a handle: its index in celt_modes (static_mode_list), 1 = mode48000_960_120; a
' const kiss_fft_state * is an index in celt_fft_states (fft_state48000_960_0 ... _3). The structs' const
' pointers to tables are copies of the tables (arrays). celt_vbModeWindow gives O3's opus_decoder.c the mode's
' window (celt_mode->window, after CELT_GET_MODE).
Option Explicit

Public Const MAX_PERIOD As Long = 1024

' kiss_fft.h
Public Const MAXFACTORS As Long = 8
' e.g. an fft of length 128 has 4 factors
' as far as kissfft is concerned
' 4*4*4*2

Public Type kiss_fft_state
    nfft As Long
    scale As Single                     ' opus_val16
    shift As Long
    factors(0 To 2 * MAXFACTORS - 1) As Long    ' opus_int16 factors[2*MAXFACTORS]
    bitrev() As Long                    ' const opus_int16 *bitrev
    twiddles_r() As Single              ' const kiss_twiddle_cpx *twiddles: (VB) .r and .i
    twiddles_i() As Single
    ' arch_fft_state *arch_fft: NULL (no OVERRIDE_FFT)
End Type

' mdct.h
Public Type mdct_lookup
    n As Long
    maxshift As Long
    kfft(0 To 3) As Long                ' const kiss_fft_state *kfft[4]: (VB) indexes into celt_fft_states
    trig() As Single                    ' const kiss_twiddle_scalar * OPUS_RESTRICT trig
End Type

' modes.h
Public Type PulseCache
    size As Long
    index() As Integer                  ' const opus_int16 *index
    bits() As Byte                      ' const unsigned char *bits
    caps() As Byte                      ' const unsigned char *caps
End Type

'* Mode definition (opaque)
' @brief Mode definition
Public Type CELTMode                    ' struct OpusCustomMode
    Fs As Long                          ' opus_int32
    overlap As Long

    nbEBands As Long
    effEBands As Long
    preemph(0 To 3) As Single           ' opus_val16
    eBands() As Integer                 '*< Definition for each "pseudo-critical band"

    maxLM As Long
    nbShortMdcts As Long
    shortMdctSize As Long

    nbAllocVectors As Long              '*< Number of lines in the matrix below
    allocVectors() As Byte              '*< Number of bits in each band for several rates
    logN() As Integer

    window() As Single                  ' const opus_val16 *window
    mdct As mdct_lookup
    cache As PulseCache
End Type

' static_modes_float.h: the states and the mode, filled by celt_modes_init
Public celt_fft_states(0 To 3) As kiss_fft_state    ' fft_state48000_960_0 ... _3
' List of all the available modes
Private Const TOTAL_MODES As Long = 1
Public celt_modes(1 To TOTAL_MODES) As CELTMode     ' static_mode_list: mode48000_960_120

Private celt_modes_done As Boolean

' (VB) 1<<n, for the C's shifts (VB's 2 ^ n is a floating point power: slow); celt_pow2(31) is the bits of 2^31
Public celt_pow2(0 To 31) As Long

' (VB) the static initialisers of static_modes_float.h
Private Sub celt_modes_init()
    Dim i As Long
    celt_pow2(0) = 1
    For i = 1 To 30
        celt_pow2(i) = celt_pow2(i - 1) * 2
    Next
    celt_pow2(31) = &H80000000
    celt_tables_init

    ' fft_state48000_960_0 ... _3: nfft, scale, shift, factors, bitrev, twiddles
    fft_state_init 0, 480, 0.002083333!, -1, fft_bitrev480, Array(5, 96, 3, 32, 4, 8, 2, 4, 4, 1)
    fft_state_init 1, 240, 0.004166667!, 1, fft_bitrev240, Array(5, 48, 3, 16, 4, 4, 4, 1)
    fft_state_init 2, 120, 0.008333333!, 2, fft_bitrev120, Array(5, 24, 3, 8, 2, 4, 4, 1)
    fft_state_init 3, 60, 0.016666667!, 3, fft_bitrev60, Array(5, 12, 3, 4, 4, 1)

    With celt_modes(1)                  ' mode48000_960_120
        .Fs = 48000                     ' Fs
        .overlap = 120                  ' overlap
        .nbEBands = 21                  ' nbEBands
        .effEBands = 21                 ' effEBands
        .preemph(0) = 0.8500061!        ' preemph {0.85000610f, 0.0000000f, 1.0000000f, 1.0000000f}
        .preemph(1) = 0!
        .preemph(2) = 1!
        .preemph(3) = 1!
        .eBands = eband5ms              ' eBands
        .maxLM = 3                      ' maxLM
        .nbShortMdcts = 8               ' nbShortMdcts
        .shortMdctSize = 120            ' shortMdctSize
        .nbAllocVectors = 11            ' nbAllocVectors
        .allocVectors = band_allocation ' allocVectors
        .logN = logN400                 ' logN
        .window = window120             ' window
        ' mdct: {1920, 3, {&fft_state48000_960_0, ..._1, ..._2, ..._3}, mdct_twiddles960}
        .mdct.n = 1920
        .mdct.maxshift = 3
        For i = 0 To 3
            .mdct.kfft(i) = i
        Next
        .mdct.trig = mdct_twiddles960
        ' cache: {392, cache_index50, cache_bits50, cache_caps50}
        .cache.size = 392
        .cache.index = cache_index50
        .cache.bits = cache_bits50
        .cache.caps = cache_caps50
    End With
    celt_modes_done = True
End Sub

' (VB) one kiss_fft_state initialiser
Private Sub fft_state_init(ByVal k As Long, ByVal nfft As Long, ByVal scale_ As Single, ByVal shift As Long, bitrev() As Integer, ByVal factors As Variant)
    Dim i As Long
    With celt_fft_states(k)
        .nfft = nfft
        .scale = scale_
        .shift = shift
        For i = 0 To 2 * MAXFACTORS - 1
            .factors(i) = 0
        Next
        For i = 0 To UBound(factors)
            .factors(i) = factors(i)
        Next
        ReDim .bitrev(0 To nfft - 1)
        For i = 0 To nfft - 1
            .bitrev(i) = bitrev(i)
        Next
        .twiddles_r = fft_twiddles48000_960_r       ' fft_twiddles48000_960
        .twiddles_i = fft_twiddles48000_960_i
    End With
End Sub

' (VB) the result is the mode's handle (0 = NULL); error_ receives *error
Public Function opus_custom_mode_create(ByVal Fs As Long, ByVal frame_size As Long, Optional error_ As Long) As Long
    Dim i As Long, j As Long
    If Not celt_modes_done Then celt_modes_init

    ' #ifndef CUSTOM_MODES_ONLY
    For i = 1 To TOTAL_MODES
        For j = 0 To 3
            If Fs = celt_modes(i).Fs And (frame_size * celt_pow2(j)) = celt_modes(i).shortMdctSize * celt_modes(i).nbShortMdcts Then
                error_ = OPUS_OK
                opus_custom_mode_create = i
                Exit Function
            End If
        Next
    Next
    ' #endif /* CUSTOM_MODES_ONLY */

    ' #ifndef CUSTOM_MODES
    error_ = OPUS_BAD_ARG
    opus_custom_mode_create = 0
    ' #else: (custom modes, not compiled)
End Function

' (VB) for O3's opus_decoder.c: mode->window and mode->overlap of the mode handle (CELT_GET_MODE's)
Public Sub celt_vbModeWindow(ByVal mode As Long, window() As Single, overlap As Long)
    If Not celt_modes_done Then celt_modes_init
    window = celt_modes(mode).window
    overlap = celt_modes(mode).overlap
End Sub
