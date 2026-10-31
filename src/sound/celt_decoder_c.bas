Attribute VB_Name = "celt_decoder_c"
' ref/opus/celt (libopus 1.5.2): celt_decoder.c, the float build without CUSTOM_MODES, ENABLE_DEEP_PLC,
' ENABLE_HARDENING or ENABLE_ASSERTIONS (cluster K, fork O2), whole but for what those leave out: the
' decoder state, deemphasis_stereo_simple, deemphasis, celt_synthesis, tf_decode, celt_plc_pitch_search,
' prefilter_and_fold, celt_decode_lost (the packet loss concealment: noise-based and pitch-based),
' celt_decode_with_ec_dred, celt_decode_with_ec, opus_custom_decoder_ctl (as celt_decoder_ctl),
' celt_decoder_get_size, celt_decoder_init. Not compiled in this build: validate_celt_decoder, the DEEP_PLC
' parts (update_plc_state, sinc_filter, the neural concealment), opus_custom_decoder_create / _destroy /
' opus_custom_decode / _float (CUSTOM_MODES), the CUSTOM_MODES branches.
' (VB) K_DESIGN.md's contract for O3's opus_decoder.c: a CELTDecoder * is a handle > 0 into this module's pool
' (celt_decoder_vbNew / celt_decoder_vbFree; the C carves it out of the OpusDecoder's memory). The state's
' _decode_mem, lpc, oldBandE, oldLogE, oldLogE2, backgroundLogE are arrays of their own (the C lays them out
' after the struct); celt_sig *decode_mem[c] and out_syn[c] are offsets in decode_mem. data NULL is
' dataIsNull, ec_dec *dec NULL is decIsNull; pcm (opus_val16 *: float) is pcm() from pcmOffset. A ctl's
' argument is value (a GET writes it, a SET reads it); CELT_GET_MODE writes the mode's handle (1). The
' constants of opus_defines.h are O3's (opus_defines_h): OPUS_OK, OPUS_BAD_ARG, OPUS_INTERNAL_ERROR,
' OPUS_UNIMPLEMENTED, OPUS_ALLOC_FAIL, OPUS_RESET_STATE, OPUS_SET_COMPLEXITY_REQUEST,
' OPUS_GET_COMPLEXITY_REQUEST, OPUS_GET_LOOKAHEAD_REQUEST, OPUS_GET_PITCH_REQUEST,
' OPUS_GET_FINAL_RANGE_REQUEST, OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST,
' OPUS_GET_PHASE_INVERSION_DISABLED_REQUEST. The celt_norm X is celt_bands_c's celt_norm_buf from 0.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Function celt_clearfp Lib "msvcrt" Alias "_clearfp" () As Long

' The maximum pitch lag to allow in the pitch-based PLC. It's possible to save
'   CPU time in the PLC pitch search by making this smaller than MAX_PERIOD. The
'   current value corresponds to a pitch of 66.67 Hz.
Private Const PLC_PITCH_LAG_MAX As Long = 720
' The minimum pitch lag to allow in the pitch-based PLC. This corresponds to a
'   pitch of 480 Hz.
Private Const PLC_PITCH_LAG_MIN As Long = 100

'*********************************************************************
'                                                                    *
'                             DECODER                                *
'                                                                    *
'*********************************************************************
Private Const DECODE_BUFFER_SIZE As Long = 2048

' arch.h's float constants
Private Const VERY_SMALL As Single = 1E-30!
Private Const SCALEOUT_ As Single = 1! / 32768!         ' SCALEOUT(a) is (a)*(1/CELT_SIG_SCALE)

'* Decoder state
' @brief Decoder state
Private Type CELTDecoder                                ' struct OpusCustomDecoder
    inUse As Boolean                                    ' (VB) the pool slot is taken
    mode As Long                                        ' const OpusCustomMode *mode: the mode's handle
    overlap As Long
    channels As Long
    stream_channels As Long

    downsample As Long
    start As Long
    end_ As Long                                        ' end
    signalling As Long
    disable_inv As Long
    complexity As Long
    arch As Long

    ' Everything beyond this point gets cleared on a reset
    ' #define DECODER_RESET_START rng

    rng As Long                                         ' opus_uint32
    error_ As Long                                      ' error
    last_pitch_index As Long
    loss_duration As Long
    skip_plc As Long
    postfilter_period As Long
    postfilter_period_old As Long
    postfilter_gain As Single
    postfilter_gain_old As Single
    postfilter_tapset As Long
    postfilter_tapset_old As Long
    prefilter_and_fold As Long

    preemph_memD(0 To 1) As Single

    decode_mem() As Single      ' _decode_mem: Size = channels*(DECODE_BUFFER_SIZE+mode->overlap)
    lpc() As Single             ' opus_val16 lpc[],  Size = channels*CELT_LPC_ORDER
    oldBandE() As Single        ' opus_val16 oldEBands[], Size = 2*mode->nbEBands
    oldLogE() As Single         ' opus_val16 oldLogE[], Size = 2*mode->nbEBands
    oldLogE2() As Single        ' opus_val16 oldLogE2[], Size = 2*mode->nbEBands
    backgroundLogE() As Single  ' opus_val16 backgroundLogE[], Size = 2*mode->nbEBands
End Type

' (VB) the pool of decoders (handle = index)
Private celt_decs() As CELTDecoder
Private celt_decs_n As Long

' (VB) the C's ALLOC'd scratch
Private synth_freq() As Single          ' celt_synthesis' freq
Private deemph_scratch() As Single      ' deemphasis' scratch
Private fold_etmp() As Single           ' prefilter_and_fold's etmp
Private plc_exc() As Single             ' celt_decode_lost's _exc (exc = _exc + CELT_LPC_ORDER)
Private plc_fir_tmp() As Single         ' celt_decode_lost's fir_tmp
Private plc_ac(0 To CELT_LPC_ORDER) As Single
Private plc_lpc_mem(0 To CELT_LPC_ORDER - 1) As Single
Private lp_pitch_buf() As Single        ' celt_plc_pitch_search's lp_pitch_buf
Private dec_tf_res() As Long
Private dec_cap() As Long
Private dec_offsets() As Long
Private dec_fine_quant() As Long
Private dec_pulses() As Long
Private dec_fine_priority() As Long
Private dec_collapse_masks() As Byte
Private celt_dec_local As ec_ctx        ' celt_decode_with_ec_dred's _dec
Private celt_decoder_done As Boolean

' (VB) test instrumentation (VB-only; tools/celt_test.py prints them): how often the decoder took its paths.
' 0 frames decoded from data, 1 transient frames, 2 anti-collapse on, 3 stereo frames with intensity stereo,
' 4 dual stereo frames, 5 silence frames, 6 post-filter on, 7 pitch-based PLC frames, 8 noise-based PLC frames,
' 9 mono streams decoded to two channels, 10 stereo streams downmixed to one, 11 downsampled frames,
' 12 spread other than SPREAD_NORMAL, 13 frames with a dynalloc boost, 14 frames with skipped (uncoded) bands
Public celt_vbCounts(0 To 15) As Long

' (VB) the other modules' tables and this module's scratch, once
Private Sub celt_decoder_globals()
    If celt_decoder_done Then Exit Sub
    celt_c_init
    celt_bands_init
    ReDim synth_freq(0 To 959)
    ReDim deemph_scratch(0 To 959)
    ReDim fold_etmp(0 To 119)
    ReDim plc_exc(0 To MAX_PERIOD + CELT_LPC_ORDER - 1)
    ReDim plc_fir_tmp(0 To MAX_PERIOD - 1)
    ReDim lp_pitch_buf(0 To DECODE_BUFFER_SIZE \ 2 - 1)
    ReDim dec_tf_res(0 To 24)
    ReDim dec_cap(0 To 24)
    ReDim dec_offsets(0 To 24)
    ReDim dec_fine_quant(0 To 24)
    ReDim dec_pulses(0 To 24)
    ReDim dec_fine_priority(0 To 24)
    ReDim dec_collapse_masks(0 To 49)
    celt_decoder_done = True
End Sub

' (VB) a new CELTDecoder (the C's memory inside the OpusDecoder): its handle
Public Function celt_decoder_vbNew() As Long
    Dim i As Long
    For i = 1 To celt_decs_n
        If Not celt_decs(i).inUse Then Exit For
    Next
    If i > celt_decs_n Then
        celt_decs_n = celt_decs_n + 1
        ReDim Preserve celt_decs(1 To celt_decs_n)
        i = celt_decs_n
    End If
    celt_decs(i).inUse = True
    celt_decoder_vbNew = i
End Function

' (VB) frees the handle's decoder
Public Sub celt_decoder_vbFree(ByVal st As Long)
    Dim blank As CELTDecoder
    If st < 1 Or st > celt_decs_n Then Exit Sub
    celt_decs(st) = blank
End Sub

' #if defined(ENABLE_HARDENING) || defined(ENABLE_ASSERTIONS): validate_celt_decoder (not compiled)

Public Function celt_decoder_get_size(ByVal channels As Long) As Long
    Dim mode As Long
    mode = opus_custom_mode_create(48000, 960)
    celt_decoder_get_size = opus_custom_decoder_get_size(mode, channels)
End Function

' (VB) the C's byte count for a 32-bit build: sizeof(struct CELTDecoder) is 104 there
Private Function opus_custom_decoder_get_size(ByVal mode As Long, ByVal channels As Long) As Long
    Dim size As Long
    size = 104 _
         + (channels * (DECODE_BUFFER_SIZE + celt_modes(mode).overlap) - 1) * 4 _
         + channels * CELT_LPC_ORDER * 4 _
         + 4 * 2 * celt_modes(mode).nbEBands * 4
    opus_custom_decoder_get_size = size
End Function

' #ifdef CUSTOM_MODES: opus_custom_decoder_create (not compiled)

Public Function celt_decoder_init(ByVal st As Long, ByVal sampling_rate As Long, ByVal channels As Long) As Long
    Dim ret As Long
    ret = opus_custom_decoder_init(st, opus_custom_mode_create(48000, 960), channels)
    If ret <> OPUS_OK Then
        celt_decoder_init = ret
        Exit Function
    End If
    celt_decs(st).downsample = resampling_factor(sampling_rate)
    If celt_decs(st).downsample = 0 Then
        celt_decoder_init = OPUS_BAD_ARG
    Else
        celt_decoder_init = OPUS_OK
    End If
End Function

Private Function opus_custom_decoder_init(ByVal st As Long, ByVal mode As Long, ByVal channels As Long) As Long
    Dim blank As CELTDecoder, n As Long
    If channels < 0 Or channels > 2 Then
        opus_custom_decoder_init = OPUS_BAD_ARG
        Exit Function
    End If

    If st < 1 Or st > celt_decs_n Then
        opus_custom_decoder_init = OPUS_ALLOC_FAIL
        Exit Function
    End If
    celt_decoder_globals

    ' OPUS_CLEAR((char*)st, opus_custom_decoder_get_size(mode, channels));
    celt_decs(st) = blank
    celt_decs(st).inUse = True
    n = channels * (DECODE_BUFFER_SIZE + celt_modes(mode).overlap)
    If n < 1 Then n = 1
    ReDim celt_decs(st).decode_mem(0 To n - 1)
    n = channels * CELT_LPC_ORDER
    If n < 1 Then n = 1
    ReDim celt_decs(st).lpc(0 To n - 1)
    n = 2 * celt_modes(mode).nbEBands
    ReDim celt_decs(st).oldBandE(0 To n - 1)
    ReDim celt_decs(st).oldLogE(0 To n - 1)
    ReDim celt_decs(st).oldLogE2(0 To n - 1)
    ReDim celt_decs(st).backgroundLogE(0 To n - 1)

    With celt_decs(st)
        .mode = mode
        .overlap = celt_modes(mode).overlap
        .channels = channels
        .stream_channels = channels

        .downsample = 1
        .start = 0
        .end_ = celt_modes(mode).effEBands
        .signalling = 1
        ' #ifndef DISABLE_UPDATE_DRAFT
        If channels = 1 Then .disable_inv = 1 Else .disable_inv = 0
        ' #else ... #endif
        .arch = 0                                       ' opus_select_arch(): no RTCD
    End With

    celt_decoder_ctl st, OPUS_RESET_STATE

    opus_custom_decoder_init = OPUS_OK
End Function

' #ifdef CUSTOM_MODES: opus_custom_decoder_destroy (not compiled)

' #ifndef CUSTOM_MODES
' Special case for stereo with no downsampling and no accumulation. This is
'   quite common and we can make it faster by processing both channels in the
'   same loop, reducing overhead due to the dependency loop in the IIR filter.
' (VB) celt_sig *in[]: x() with the channels at x0o and x1o; pcm: pcm() from pcmo
Private Sub deemphasis_stereo_simple(x() As Single, ByVal x0o As Long, ByVal x1o As Long, pcm() As Single, ByVal pcmo As Long, ByVal N As Long, ByVal coef0 As Single, mem() As Single)
    Dim m0 As Single, m1 As Single, j As Long, tmp0 As Single, tmp1 As Single
    m0 = mem(0)
    m1 = mem(1)
    For j = 0 To N - 1
        ' Add VERY_SMALL to x[] first to reduce dependency chain.
        tmp0 = x(x0o + j) + VERY_SMALL + m0
        tmp1 = x(x1o + j) + VERY_SMALL + m1
        m0 = coef0 * tmp0                               ' MULT16_32_Q15(coef0, tmp0)
        m1 = coef0 * tmp1
        pcm(pcmo + 2 * j) = tmp0 * SCALEOUT_            ' SCALEOUT(SIG2WORD16(tmp0))
        pcm(pcmo + 2 * j + 1) = tmp1 * SCALEOUT_
    Next
    mem(0) = m0
    mem(1) = m1
End Sub
' #endif

' (VB) celt_sig *in[]: x() with the channels at x0o and x1o; pcm: pcm() from pcmo; coef: coef() from 0
Private Sub deemphasis(x() As Single, ByVal x0o As Long, ByVal x1o As Long, pcm() As Single, ByVal pcmo As Long, ByVal N As Long, ByVal C As Long, ByVal downsample As Long, coef() As Single, _
        mem() As Single, ByVal accum As Long)
    Dim c_ As Long, Nd As Long, apply_downsampling As Long, coef0 As Single, j As Long, m As Single, tmp As Single
    Dim xo As Long, y As Long
    ' #ifndef CUSTOM_MODES
    ' Short version for common case.
    If downsample = 1 And C = 2 And accum = 0 Then
        deemphasis_stereo_simple x, x0o, x1o, pcm, pcmo, N, coef(0), mem
        Exit Sub
    End If
    ' #endif
    ' #ifndef FIXED_POINT
    ' (void)accum;
    ' celt_assert(accum==0);
    ' #endif
    ' ALLOC(scratch, N, celt_sig);
    coef0 = coef(0)
    Nd = N \ downsample
    c_ = 0
    Do
        m = mem(c_)
        If c_ = 0 Then xo = x0o Else xo = x1o           ' x =in[c];
        y = pcmo + c_                                   ' y = pcm+c;
        ' #ifdef CUSTOM_MODES ... #endif
        If downsample > 1 Then
            ' Shortcut for the standard (non-custom modes) case
            For j = 0 To N - 1
                tmp = x(xo + j) + VERY_SMALL + m
                m = coef0 * tmp
                deemph_scratch(j) = tmp
            Next
            apply_downsampling = 1
        Else
            ' Shortcut for the standard (non-custom modes) case
            ' #ifdef FIXED_POINT (accum) ... #endif
            For j = 0 To N - 1
                tmp = x(xo + j) + VERY_SMALL + m
                m = coef0 * tmp
                pcm(y + j * C) = tmp * SCALEOUT_
            Next
        End If
        mem(c_) = m

        If apply_downsampling <> 0 Then
            ' Perform down-sampling
            ' #ifdef FIXED_POINT (accum) ... #endif
            For j = 0 To Nd - 1
                pcm(y + j * C) = deemph_scratch(j * downsample) * SCALEOUT_
            Next
        End If
        c_ = c_ + 1
    Loop While c_ < C
End Sub

' (VB) X: celt_norm_buf from X; celt_sig * out_syn[]: the decoder memory dm() and the channels' offsets os0, os1;
' oldBandE: oldBandE() from 0
Private Sub celt_synthesis(ByVal mode As Long, ByVal X As Long, dm() As Single, ByVal os0 As Long, ByVal os1 As Long, _
        oldBandE() As Single, ByVal start As Long, ByVal effEnd As Long, ByVal C As Long, ByVal CC As Long, _
        ByVal isTransient As Long, ByVal LM As Long, ByVal downsample As Long, _
        ByVal silence As Long, ByVal arch As Long)
    Dim c_ As Long, i As Long, M_ As Long, b As Long, B_ As Long, N As Long, NB As Long, shift As Long
    Dim nbEBands As Long, overlap As Long, freq2 As Long, os As Long

    overlap = celt_modes(mode).overlap
    nbEBands = celt_modes(mode).nbEBands
    N = celt_modes(mode).shortMdctSize * celt_pow2(LM)
    ' ALLOC(freq, N, celt_sig); /**< Interleaved signal MDCTs */
    M_ = celt_pow2(LM)

    If isTransient <> 0 Then
        B_ = M_
        NB = celt_modes(mode).shortMdctSize
        shift = celt_modes(mode).maxLM
    Else
        B_ = 1
        NB = celt_modes(mode).shortMdctSize * celt_pow2(LM)
        shift = celt_modes(mode).maxLM - LM
    End If

    If CC = 2 And C = 1 Then
        celt_vbCounts(9) = celt_vbCounts(9) + 1
        ' Copying a mono streams to two channels
        denormalise_bands mode, X, synth_freq, 0, oldBandE, 0, start, effEnd, M_, _
              downsample, silence
        ' Store a temporary copy in the output buffer because the IMDCT destroys its input.
        freq2 = os1 + overlap \ 2
        CopyMemory dm(freq2), synth_freq(0), N * 4     ' OPUS_COPY(freq2, freq, N);
        For b = 0 To B_ - 1
            clt_mdct_backward_c celt_modes(mode).mdct, dm, freq2 + b, dm, os0 + NB * b, celt_modes(mode).window, overlap, shift, B_, arch
        Next
        For b = 0 To B_ - 1
            clt_mdct_backward_c celt_modes(mode).mdct, synth_freq, b, dm, os1 + NB * b, celt_modes(mode).window, overlap, shift, B_, arch
        Next
    ElseIf CC = 1 And C = 2 Then
        celt_vbCounts(10) = celt_vbCounts(10) + 1
        ' Downmixing a stereo stream to mono
        freq2 = os0 + overlap \ 2
        denormalise_bands mode, X, synth_freq, 0, oldBandE, 0, start, effEnd, M_, _
              downsample, silence
        ' Use the output buffer as temp array before downmixing.
        denormalise_bands mode, X + N, dm, freq2, oldBandE, nbEBands, start, effEnd, M_, _
              downsample, silence
        For i = 0 To N - 1
            synth_freq(i) = 0.5! * synth_freq(i) + 0.5! * dm(freq2 + i)    ' ADD32(HALF32(freq[i]), HALF32(freq2[i]))
        Next
        For b = 0 To B_ - 1
            clt_mdct_backward_c celt_modes(mode).mdct, synth_freq, b, dm, os0 + NB * b, celt_modes(mode).window, overlap, shift, B_, arch
        Next
    Else
        ' Normal case (mono or stereo)
        c_ = 0
        Do
            If c_ = 0 Then os = os0 Else os = os1
            denormalise_bands mode, X + c_ * N, synth_freq, 0, oldBandE, c_ * nbEBands, start, effEnd, M_, _
                  downsample, silence
            For b = 0 To B_ - 1
                clt_mdct_backward_c celt_modes(mode).mdct, synth_freq, b, dm, os + NB * b, celt_modes(mode).window, overlap, shift, B_, arch
            Next
            c_ = c_ + 1
        Loop While c_ < CC
    End If
    ' Saturate IMDCT output so that we can't overflow in the pitch postfilter
    '   or in the
    ' (VB) out_syn[c][i] = SATURATE(out_syn[c][i], SIG_SAT): nothing in the float build
End Sub

' (VB) tf_res: tf_res() from 0
Private Sub tf_decode(ByVal start As Long, ByVal end_ As Long, ByVal isTransient As Long, tf_res() As Long, ByVal LM As Long, dec As ec_ctx)
    Dim i As Long, curr As Long, tf_select As Long, tf_select_rsv As Long, tf_changed As Long, logp As Long
    Dim budget As Long, tell As Long

    budget = dec.storage * 8
    tell = ec_tell(dec)
    If isTransient <> 0 Then logp = 2 Else logp = 4
    If LM > 0 And tell + logp + 1 <= budget Then tf_select_rsv = 1 Else tf_select_rsv = 0
    budget = budget - tf_select_rsv
    tf_changed = 0: curr = 0
    For i = start To end_ - 1
        If tell + logp <= budget Then
            curr = curr Xor ec_dec_bit_logp(dec, logp)
            tell = ec_tell(dec)
            tf_changed = tf_changed Or curr
        End If
        tf_res(i) = curr
        If isTransient <> 0 Then logp = 4 Else logp = 5
    Next
    tf_select = 0
    If tf_select_rsv <> 0 Then
        If tf_select_table(LM * 8 + 4 * isTransient + 0 + tf_changed) <> tf_select_table(LM * 8 + 4 * isTransient + 2 + tf_changed) Then
            tf_select = ec_dec_bit_logp(dec, 1)
        End If
    End If
    For i = start To end_ - 1
        tf_res(i) = tf_select_table(LM * 8 + 4 * isTransient + 2 * tf_select + tf_res(i))
    Next
End Sub

' (VB) celt_sig *decode_mem[2]: dm() with the channels at dm0 and dm1
Private Function celt_plc_pitch_search(dm() As Single, ByVal dm0 As Long, ByVal dm1 As Long, ByVal C As Long, ByVal arch As Long) As Long
    Dim pitch_index As Long
    ' ALLOC( lp_pitch_buf, DECODE_BUFFER_SIZE>>1, opus_val16 );
    pitch_downsample dm, dm0, dm1, lp_pitch_buf, _
          DECODE_BUFFER_SIZE, C, arch
    pitch_search lp_pitch_buf, PLC_PITCH_LAG_MAX \ 2, lp_pitch_buf, 0, _
          DECODE_BUFFER_SIZE - PLC_PITCH_LAG_MAX, _
          PLC_PITCH_LAG_MAX - PLC_PITCH_LAG_MIN, pitch_index, arch
    pitch_index = PLC_PITCH_LAG_MAX - pitch_index
    celt_plc_pitch_search = pitch_index
End Function

Private Sub prefilter_and_fold(st As CELTDecoder, ByVal N As Long)
    Dim c_ As Long, CC As Long, i As Long, overlap As Long, dm As Long
    ' const OpusCustomMode *mode = st->mode;
    overlap = st.overlap
    CC = st.channels
    ' ALLOC(etmp, overlap, opus_val32);
    ' decode_mem[c] = st->_decode_mem + c*(DECODE_BUFFER_SIZE+overlap);

    c_ = 0
    Do
        dm = c_ * (DECODE_BUFFER_SIZE + overlap)
        ' Apply the pre-filter to the MDCT overlap for the next frame because
        '   the post-filter will be re-applied in the decoder after the MDCT
        '   overlap.
        comb_filter fold_etmp, 0, st.decode_mem, dm + DECODE_BUFFER_SIZE - N, _
           st.postfilter_period_old, st.postfilter_period, overlap, _
           -st.postfilter_gain_old, -st.postfilter_gain, _
           st.postfilter_tapset_old, st.postfilter_tapset, celt_modes(st.mode).window, 0, st.arch

        ' Simulate TDAC on the concealed audio so that it blends with the
        '   MDCT of the next frame.
        With celt_modes(st.mode)
            For i = 0 To overlap \ 2 - 1
                st.decode_mem(dm + DECODE_BUFFER_SIZE - N + i) = _
                   .window(i) * fold_etmp(overlap - 1 - i) _
                   + .window(overlap - i - 1) * fold_etmp(i)        ' MULT16_32_Q15(mode->window[i], etmp[overlap-1-i]) + ...
            Next
        End With
        c_ = c_ + 1
    Loop While c_ < CC
End Sub

' #ifdef ENABLE_DEEP_PLC: sinc_filter, update_plc_state (not compiled)

Private Sub celt_decode_lost(st As CELTDecoder, ByVal N As Long, ByVal LM As Long)
    Dim c_ As Long, i As Long, j As Long, C As Long, mode As Long, nbEBands As Long, overlap As Long
    Dim start As Long, loss_duration As Long, noise_based As Long, nb As Long
    Dim dm0 As Long, dm1 As Long, os0 As Long, os1 As Long
    Dim seed As Long, end_ As Long, effEnd As Long, decay As Single, boffs As Long, blen As Long, k As Long
    Dim exc_length As Long, fade As Single, pitch_index As Long, attenuation As Single, S1 As Single, S2 As Single
    Dim buf As Long, extrapolation_offset As Long, extrapolation_len As Long, E1 As Single, E2 As Single
    Dim decay_length As Long, e As Single, tmp As Single, ratio As Single, tmp_g As Single
    Dim sd As Double, s As Long, explosion As Boolean

    C = st.channels
    mode = st.mode
    nbEBands = celt_modes(mode).nbEBands
    overlap = celt_modes(mode).overlap
    ' eBands = mode->eBands;

    ' decode_mem[c] = st->_decode_mem + c*(DECODE_BUFFER_SIZE+overlap); out_syn[c] = decode_mem[c]+DECODE_BUFFER_SIZE-N;
    dm0 = 0
    dm1 = DECODE_BUFFER_SIZE + overlap
    os0 = dm0 + DECODE_BUFFER_SIZE - N
    os1 = dm1 + DECODE_BUFFER_SIZE - N
    ' lpc, oldBandE, oldLogE, oldLogE2, backgroundLogE: (VB) st's arrays

    loss_duration = st.loss_duration
    start = st.start
    ' #ifdef ENABLE_DEEP_PLC ... #else
    If loss_duration >= 40 Or start <> 0 Or st.skip_plc <> 0 Then noise_based = 1
    ' #endif
    If noise_based <> 0 Then
        ' Noise-based PLC/CNG
        end_ = st.end_
        effEnd = end_
        If celt_modes(mode).effEBands < effEnd Then effEnd = celt_modes(mode).effEBands
        If effEnd < start Then effEnd = start           ' IMAX(start, IMIN(end, mode->effEBands))

        ' ALLOC(X, C*N, celt_norm);   /**< Interleaved normalised MDCTs */  (VB) celt_norm_buf from 0
        c_ = 0
        Do
            k = c_ * (DECODE_BUFFER_SIZE + overlap)
            CopyMemory st.decode_mem(k), st.decode_mem(k + N), (DECODE_BUFFER_SIZE - N + overlap) * 4  ' OPUS_MOVE
            c_ = c_ + 1
        Loop While c_ < C

        If st.prefilter_and_fold <> 0 Then
            prefilter_and_fold st, N
        End If

        ' Energy decay
        If loss_duration = 0 Then decay = 1.5! Else decay = 0.5!   ' QCONST16(1.5f, DB_SHIFT) : QCONST16(.5f, DB_SHIFT)
        c_ = 0
        Do
            For i = start To end_ - 1
                nb = c_ * nbEBands + i
                ' MAX16(backgroundLogE[c*nbEBands+i], oldBandE[c*nbEBands+i] - decay)
                If st.backgroundLogE(nb) > st.oldBandE(nb) - decay Then
                    st.oldBandE(nb) = st.backgroundLogE(nb)
                Else
                    st.oldBandE(nb) = st.oldBandE(nb) - decay
                End If
            Next
            c_ = c_ + 1
        Loop While c_ < C
        seed = st.rng
        For c_ = 0 To C - 1
            For i = start To effEnd - 1
                boffs = N * c_ + celt_modes(mode).eBands(i) * celt_pow2(LM)
                blen = (celt_modes(mode).eBands(i + 1) - celt_modes(mode).eBands(i)) * celt_pow2(LM)
                ' (VB) seed = celt_lcg_rand(seed), written out with the seed in a Double
                sd = seed
                If sd < 0 Then sd = sd + 4294967296#
                For j = 0 To blen - 1
                    sd = 1664525# * sd + 1013904223#
                    sd = sd - Int(sd / 4294967296#) * 4294967296#
                    If sd >= 2147483648# Then s = sd - 4294967296# Else s = sd
                    celt_norm_buf(boffs + j) = (s And &HFFF00000) \ &H100000        ' (celt_norm)((opus_int32)seed>>20)
                Next
                seed = s
                renormalise_vector celt_norm_buf, boffs, blen, 1!, st.arch
            Next
        Next
        st.rng = seed
        celt_vbCounts(8) = celt_vbCounts(8) + 1

        celt_synthesis mode, 0, st.decode_mem, os0, os1, st.oldBandE, start, effEnd, C, C, 0, LM, st.downsample, 0, st.arch
        st.prefilter_and_fold = 0
        ' Skip regular PLC until we get two consecutive packets.
        st.skip_plc = 1
    Else
        ' Pitch-based PLC
        fade = 1!                                       ' Q15ONE

        If loss_duration = 0 Then
            ' #ifdef ENABLE_DEEP_PLC ... #endif
            pitch_index = celt_plc_pitch_search(st.decode_mem, dm0, dm1, C, st.arch)
            st.last_pitch_index = pitch_index
        Else
            pitch_index = st.last_pitch_index
            fade = 0.8!                                 ' QCONST16(.8f,15)
        End If

        ' We want the excitation for 2 pitch periods in order to look for a
        '   decaying signal, but we can't get more than MAX_PERIOD.
        exc_length = 2 * pitch_index
        If exc_length > MAX_PERIOD Then exc_length = MAX_PERIOD

        ' ALLOC(_exc, MAX_PERIOD+CELT_LPC_ORDER, opus_val16); ALLOC(fir_tmp, exc_length, opus_val16);
        ' exc = _exc+CELT_LPC_ORDER: (VB) exc[i] is plc_exc(CELT_LPC_ORDER + i)
        ' window = mode->window;
        c_ = 0
        Do
            S1 = 0

            buf = c_ * (DECODE_BUFFER_SIZE + overlap)   ' buf = decode_mem[c];
            For i = 0 To MAX_PERIOD + CELT_LPC_ORDER - 1
                plc_exc(i) = st.decode_mem(buf + DECODE_BUFFER_SIZE - MAX_PERIOD - CELT_LPC_ORDER + i)  ' SROUND16(..., SIG_SHIFT)
            Next

            If loss_duration = 0 Then
                ' Compute LPC coefficients for the last MAX_PERIOD samples before
                '   the first loss so we can work in the excitation-filter domain.
                celt_autocorr_ plc_exc, CELT_LPC_ORDER, plc_ac, celt_modes(mode).window, overlap, _
                       CELT_LPC_ORDER, MAX_PERIOD, st.arch
                ' Add a noise floor of -40 dB.
                ' #ifdef FIXED_POINT ... #else
                plc_ac(0) = plc_ac(0) * 1.0001!
                ' #endif
                ' Use lag windowing to stabilize the Levinson-Durbin recursion.
                For i = 1 To CELT_LPC_ORDER
                    'ac[i] *= exp(-.5*(2*M_PI*.002*i)*(2*M_PI*.002*i));
                    ' #ifdef FIXED_POINT ... #else
                    plc_ac(i) = plc_ac(i) - plc_ac(i) * (0.008! * 0.008!) * i * i
                    ' #endif
                Next
                celt_lpc_ st.lpc, c_ * CELT_LPC_ORDER, plc_ac, CELT_LPC_ORDER
                ' #ifdef FIXED_POINT: the bandwidth expansion (not compiled)
            End If
            ' Initialize the LPC history with the samples just before the start
            '   of the region for which we're computing the excitation.
            ' Compute the excitation for exc_length samples before the loss. We need the copy
            '   because celt_fir() cannot filter in-place.
            celt_fir_c plc_exc, CELT_LPC_ORDER + MAX_PERIOD - exc_length, st.lpc, c_ * CELT_LPC_ORDER, _
                  plc_fir_tmp, 0, exc_length, CELT_LPC_ORDER, st.arch
            CopyMemory plc_exc(CELT_LPC_ORDER + MAX_PERIOD - exc_length), plc_fir_tmp(0), exc_length * 4   ' OPUS_COPY

            ' Check if the waveform is decaying, and if so how fast.
            '   We do this to avoid adding energy when concealing in a segment
            '   with decaying energy.
            E1 = 1: E2 = 1
            ' #ifdef FIXED_POINT ... #endif
            decay_length = exc_length \ 2
            For i = 0 To decay_length - 1
                e = plc_exc(CELT_LPC_ORDER + MAX_PERIOD - decay_length + i)
                E1 = E1 + e * e                         ' SHR32(MULT16_16(e, e), shift)
                e = plc_exc(CELT_LPC_ORDER + MAX_PERIOD - 2 * decay_length + i)
                E2 = E2 + e * e
            Next
            If E2 < E1 Then E1 = E2                     ' MIN32(E1, E2)
            decay = Sqr(E1 / E2)                        ' celt_sqrt(frac_div32(SHR32(E1, 1), E2))

            ' Move the decoder memory one frame to the left to give us room to
            '   add the data for the new frame. We ignore the overlap that extends
            '   past the end of the buffer, because we aren't going to use it.
            CopyMemory st.decode_mem(buf), st.decode_mem(buf + N), (DECODE_BUFFER_SIZE - N) * 4     ' OPUS_MOVE

            ' Extrapolate from the end of the excitation with a period of
            '   "pitch_index", scaling down each period by an additional factor of
            '   "decay".
            extrapolation_offset = MAX_PERIOD - pitch_index
            ' We need to extrapolate enough samples to cover a complete MDCT
            '   window (including overlap/2 samples on both sides).
            extrapolation_len = N + overlap
            ' We also apply fading if this is not the first loss.
            attenuation = fade * decay                  ' MULT16_16_Q15(fade, decay)
            j = 0
            For i = 0 To extrapolation_len - 1
                If j >= pitch_index Then
                    j = j - pitch_index
                    attenuation = attenuation * decay
                End If
                st.decode_mem(buf + DECODE_BUFFER_SIZE - N + i) = _
                      attenuation * plc_exc(CELT_LPC_ORDER + extrapolation_offset + j)     ' SHL32(EXTEND32(MULT16_16_Q15(attenuation, exc[extrapolation_offset+j])), SIG_SHIFT)
                ' Compute the energy of the previously decoded signal whose
                '   excitation we're copying.
                tmp = st.decode_mem(buf + DECODE_BUFFER_SIZE - MAX_PERIOD - N + extrapolation_offset + j)
                S1 = S1 + tmp * tmp                     ' SHR32(MULT16_16(tmp, tmp), 10)
                j = j + 1
            Next
            ' Copy the last decoded samples (prior to the overlap region) to
            '   synthesis filter memory so we can have a continuous signal.
            For i = 0 To CELT_LPC_ORDER - 1
                plc_lpc_mem(i) = st.decode_mem(buf + DECODE_BUFFER_SIZE - N - 1 - i)   ' SROUND16(..., SIG_SHIFT)
            Next
            ' Apply the synthesis filter to convert the excitation back into
            '   the signal domain.
            celt_iir st.decode_mem, buf + DECODE_BUFFER_SIZE - N, st.lpc, c_ * CELT_LPC_ORDER, _
                  st.decode_mem, buf + DECODE_BUFFER_SIZE - N, extrapolation_len, CELT_LPC_ORDER, _
                  plc_lpc_mem, 0, st.arch
            ' #ifdef FIXED_POINT ... #endif

            ' Check if the synthesis energy is higher than expected, which can
            '   happen with the signal changes during our window. If so,
            '   attenuate.
            S2 = 0
            For i = 0 To extrapolation_len - 1
                tmp = st.decode_mem(buf + DECODE_BUFFER_SIZE - N + i)   ' SROUND16(..., SIG_SHIFT)
                S2 = S2 + tmp * tmp                     ' SHR32(MULT16_16(tmp, tmp), 10)
            Next
            ' This checks for an "explosion" in the synthesis.
            ' #ifdef FIXED_POINT ... #else
            ' The float test is written this way to catch NaNs in the output
            '   of the IIR filter at the same time.
            ' (VB) VB's compiled comparisons do not give C's answers for a NaN: tested first; and the NaN's
            ' sticky x87 "invalid" flag would fail VB's next float to integer conversion: cleared
            explosion = plc_isnan(S1)
            If plc_isnan(S2) Then explosion = True
            If Not explosion Then explosion = Not (S1 > 0.2! * S2)
            If explosion Then
            ' #endif
                For i = 0 To extrapolation_len - 1
                    st.decode_mem(buf + DECODE_BUFFER_SIZE - N + i) = 0
                Next
                celt_clearfp
            ElseIf S1 < S2 Then
                ratio = Sqr((S1 + 1) / (S2 + 1))        ' celt_sqrt(frac_div32(SHR32(S1,1)+1,S2+1))
                With celt_modes(mode)
                    For i = 0 To overlap - 1
                        tmp_g = 1! - .window(i) * (1! - ratio)  ' Q15ONE - MULT16_16_Q15(window[i], Q15ONE-ratio)
                        st.decode_mem(buf + DECODE_BUFFER_SIZE - N + i) = _
                              tmp_g * st.decode_mem(buf + DECODE_BUFFER_SIZE - N + i)
                    Next
                End With
                For i = overlap To extrapolation_len - 1
                    st.decode_mem(buf + DECODE_BUFFER_SIZE - N + i) = _
                          ratio * st.decode_mem(buf + DECODE_BUFFER_SIZE - N + i)
                Next
            End If

            c_ = c_ + 1
        Loop While c_ < C

        ' #ifdef ENABLE_DEEP_PLC: the neural concealment (not compiled)
        st.prefilter_and_fold = 1
        celt_vbCounts(7) = celt_vbCounts(7) + 1
    End If

    ' Saturate to soemthing large to avoid wrap-around.
    st.loss_duration = loss_duration + celt_pow2(LM)
    If st.loss_duration > 10000 Then st.loss_duration = 10000   ' IMIN(10000, loss_duration+(1<<LM))
End Sub

' (VB) whether f is a NaN (its bits)
Private Function plc_isnan(ByVal f As Single) As Boolean
    Dim bits As Long
    CopyMemory bits, f, 4
    plc_isnan = ((bits And &H7F800000) = &H7F800000) And ((bits And &H7FFFFF) <> 0)
End Function

' (VB) data NULL <=> dataIsNull; ec_dec *dec NULL <=> decIsNull; pcm: pcm() from pcmOffset
Public Function celt_decode_with_ec_dred(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, _
        ByVal dataIsNull As Boolean, ByVal len_ As Long, pcm() As Single, ByVal pcmOffset As Long, _
        ByVal frame_size As Long, dec As ec_ctx, ByVal decIsNull As Boolean, ByVal accum As Long) As Long
    ' #ifdef ENABLE_DEEP_PLC ,LPCNetPLCState *lpcnet #endif: (not compiled)
    celt_decode_with_ec_dred = decode_with_ec_(celt_decs(st), data, dataOffset, dataIsNull, len_, pcm, pcmOffset, frame_size, dec, decIsNull, accum)
End Function

' (VB) celt_decode_with_ec_dred up to its `if (dec == NULL)`, with the decoder's state ByRef
Private Function decode_with_ec_(st As CELTDecoder, data() As Byte, ByVal dataOffset As Long, _
        ByVal dataIsNull As Boolean, ByVal len_ As Long, pcm() As Single, ByVal pcmOffset As Long, _
        ByVal frame_size As Long, dec As ec_ctx, ByVal decIsNull As Boolean, ByVal accum As Long) As Long
    Dim N As Long, LM As Long, M_ As Long, CC As Long, mode As Long, overlap As Long, os0 As Long, os1 As Long

    ' VALIDATE_CELT_DECODER(st); (nothing: no ENABLE_HARDENING / ENABLE_ASSERTIONS)
    CC = st.channels
    mode = st.mode
    overlap = celt_modes(mode).overlap
    frame_size = frame_size * st.downsample

    ' lpc, oldBandE, oldLogE, oldLogE2, backgroundLogE: (VB) st's arrays

    ' #ifdef CUSTOM_MODES (signalling) ... #else
    For LM = 0 To celt_modes(mode).maxLM
        If celt_modes(mode).shortMdctSize * celt_pow2(LM) = frame_size Then Exit For
    Next
    If LM > celt_modes(mode).maxLM Then
        decode_with_ec_ = OPUS_BAD_ARG
        Exit Function
    End If
    M_ = celt_pow2(LM)

    If len_ < 0 Or len_ > 1275 Then                     ' || pcm==NULL: (VB) never
        decode_with_ec_ = OPUS_BAD_ARG
        Exit Function
    End If

    N = M_ * celt_modes(mode).shortMdctSize
    ' decode_mem[c] = st->_decode_mem + c*(DECODE_BUFFER_SIZE+overlap); out_syn[c] = decode_mem[c]+DECODE_BUFFER_SIZE-N;
    os0 = DECODE_BUFFER_SIZE - N
    os1 = (DECODE_BUFFER_SIZE + overlap) + DECODE_BUFFER_SIZE - N

    If dataIsNull Or len_ <= 1 Then
        celt_decode_lost st, N, LM
        deemphasis st.decode_mem, os0, os1, pcm, pcmOffset, N, CC, st.downsample, celt_modes(mode).preemph, st.preemph_memD, accum
        decode_with_ec_ = frame_size \ st.downsample
        Exit Function
    End If
    ' #ifdef ENABLE_DEEP_PLC ... #endif

    ' Check if there are at least two packets received consecutively before
    '  turning on the pitch-based PLC
    If st.loss_duration = 0 Then st.skip_plc = 0

    If decIsNull Then
        ec_dec_init celt_dec_local, data, dataOffset, len_
        decode_with_ec_ = decode_body_(st, len_, pcm, pcmOffset, frame_size, N, LM, M_, os0, os1, celt_dec_local, accum)
    Else
        decode_with_ec_ = decode_body_(st, len_, pcm, pcmOffset, frame_size, N, LM, M_, os0, os1, dec, accum)
    End If
End Function

' (VB) the rest of celt_decode_with_ec_dred, after `dec = &_dec;`
Private Function decode_body_(st As CELTDecoder, ByVal len_ As Long, pcm() As Single, ByVal pcmOffset As Long, _
        ByVal frame_size As Long, ByVal N As Long, ByVal LM As Long, ByVal M_ As Long, ByVal os0 As Long, ByVal os1 As Long, _
        dec As ec_ctx, ByVal accum As Long) As Long
    Dim c_ As Long, i As Long, spread_decision As Long, bits As Long
    Dim shortBlocks As Long, isTransient As Long, intra_ener As Long, CC As Long, start As Long, end_ As Long, effEnd As Long
    Dim codedBands As Long, alloc_trim As Long, postfilter_pitch As Long, postfilter_gain As Single
    Dim intensity As Long, dual_stereo As Long, total_bits As Long, balance As Long, tell As Long
    Dim dynalloc_logp As Long, postfilter_tapset As Long, anti_collapse_rsv As Long, anti_collapse_on As Long
    Dim silence As Long, C As Long, mode As Long, nbEBands As Long, overlap As Long
    Dim max_background_increase As Single, qg As Long, octave As Long
    Dim safety As Single, missing As Long, slope As Single, E0 As Single, E1 As Single, E2 As Single, k As Long
    Dim width As Long, quanta As Long, dynalloc_loop_logp As Long, boost As Long, flag As Long, os As Long, Y As Long

    CC = st.channels
    C = st.stream_channels
    mode = st.mode
    nbEBands = celt_modes(mode).nbEBands
    overlap = celt_modes(mode).overlap
    start = st.start
    end_ = st.end_

    effEnd = end_
    If effEnd > celt_modes(mode).effEBands Then effEnd = celt_modes(mode).effEBands

    If C = 1 Then
        For i = 0 To nbEBands - 1
            If st.oldBandE(nbEBands + i) > st.oldBandE(i) Then st.oldBandE(i) = st.oldBandE(nbEBands + i)  ' MAX16(oldBandE[i],oldBandE[nbEBands+i])
        Next
    End If

    total_bits = len_ * 8
    tell = ec_tell(dec)

    If tell >= total_bits Then
        silence = 1
    ElseIf tell = 1 Then
        silence = ec_dec_bit_logp(dec, 15)
    Else
        silence = 0
    End If
    If silence <> 0 Then
        ' Pretend we've read all the remaining bits
        tell = len_ * 8
        dec.nbits_total = dec.nbits_total + tell - ec_tell(dec)
    End If

    postfilter_gain = 0
    postfilter_pitch = 0
    postfilter_tapset = 0
    If start = 0 And tell + 16 <= total_bits Then
        If ec_dec_bit_logp(dec, 1) <> 0 Then
            octave = ec_dec_uint(dec, 6)
            postfilter_pitch = 16 * celt_pow2(octave) + ec_dec_bits(dec, 4 + octave) - 1
            qg = ec_dec_bits(dec, 3)
            If ec_tell(dec) + 2 <= total_bits Then
                postfilter_tapset = ec_dec_icdf(dec, tapset_icdf, 0, 2)
            End If
            postfilter_gain = 0.09375! * (qg + 1)       ' QCONST16(.09375f,15)*(qg+1)
            celt_vbCounts(6) = celt_vbCounts(6) + 1
        End If
        tell = ec_tell(dec)
    End If

    If LM > 0 And tell + 3 <= total_bits Then
        isTransient = ec_dec_bit_logp(dec, 3)
        tell = ec_tell(dec)
    Else
        isTransient = 0
    End If

    If isTransient <> 0 Then
        shortBlocks = M_
    Else
        shortBlocks = 0
    End If

    ' Decode the global flags (first symbols in the stream)
    If tell + 3 <= total_bits Then intra_ener = ec_dec_bit_logp(dec, 3) Else intra_ener = 0
    ' If recovering from packet loss, make sure we make the energy prediction safe to reduce the
    '   risk of getting loud artifacts.
    If intra_ener = 0 And st.loss_duration <> 0 Then
        c_ = 0
        Do
            safety = 0
            missing = st.loss_duration \ celt_pow2(LM)
            If missing > 10 Then missing = 10           ' IMIN(10, st->loss_duration>>LM)
            If LM = 0 Then
                safety = 1.5!                           ' QCONST16(1.5f,DB_SHIFT)
            ElseIf LM = 1 Then
                safety = 0.5!
            End If
            For i = start To end_ - 1
                k = c_ * nbEBands + i
                E1 = st.oldLogE(k)
                If st.oldLogE2(k) > E1 Then E1 = st.oldLogE2(k)
                If st.oldBandE(k) < E1 Then             ' oldBandE < MAX16(oldLogE, oldLogE2)
                    ' If energy is going down already, continue the trend.
                    E0 = st.oldBandE(k)
                    E1 = st.oldLogE(k)
                    E2 = st.oldLogE2(k)
                    slope = E1 - E0
                    If 0.5! * (E2 - E0) > slope Then slope = 0.5! * (E2 - E0)  ' MAX32(E1 - E0, HALF32(E2 - E0))
                    If (1 + missing) * slope > 0 Then E0 = E0 - (1 + missing) * slope  ' E0 -= MAX32(0, (1+missing)*slope)
                    If E0 < -20! Then E0 = -20!         ' MAX32(-QCONST16(20.f,DB_SHIFT), E0)
                    st.oldBandE(k) = E0
                Else
                    ' Otherwise take the min of the last frames.
                    If st.oldLogE(k) < st.oldBandE(k) Then st.oldBandE(k) = st.oldLogE(k)
                    If st.oldLogE2(k) < st.oldBandE(k) Then st.oldBandE(k) = st.oldLogE2(k)
                End If
                ' Shorter frames have more natural fluctuations -- play it safe.
                st.oldBandE(k) = st.oldBandE(k) - safety
            Next
            c_ = c_ + 1
        Loop While c_ < 2
    End If
    ' Get band energies
    unquant_coarse_energy mode, start, end_, st.oldBandE, _
          intra_ener, dec, C, LM

    ' ALLOC(tf_res, nbEBands, int);
    tf_decode start, end_, isTransient, dec_tf_res, LM, dec

    tell = ec_tell(dec)
    spread_decision = SPREAD_NORMAL
    If tell + 4 <= total_bits Then
        spread_decision = ec_dec_icdf(dec, spread_icdf, 0, 5)
    End If
    If spread_decision <> SPREAD_NORMAL Then celt_vbCounts(12) = celt_vbCounts(12) + 1

    ' ALLOC(cap, nbEBands, int);

    init_caps mode, dec_cap, LM, C

    ' ALLOC(offsets, nbEBands, int);

    dynalloc_logp = 6
    total_bits = total_bits * 8                         ' <<=BITRES
    tell = ec_tell_frac(dec)
    For i = start To end_ - 1
        width = C * (celt_modes(mode).eBands(i + 1) - celt_modes(mode).eBands(i)) * celt_pow2(LM)
        ' quanta is 6 bits, but no more than 1 bit/sample
        '   and no less than 1/8 bit/sample
        quanta = width                                  ' IMIN(width<<BITRES, IMAX(6<<BITRES, width))
        If quanta < 6 * 8 Then quanta = 6 * 8
        If width * 8 < quanta Then quanta = width * 8
        dynalloc_loop_logp = dynalloc_logp
        boost = 0
        Do While tell + dynalloc_loop_logp * 8 < total_bits And boost < dec_cap(i)
            flag = ec_dec_bit_logp(dec, dynalloc_loop_logp)
            tell = ec_tell_frac(dec)
            If flag = 0 Then Exit Do
            boost = boost + quanta
            total_bits = total_bits - quanta
            dynalloc_loop_logp = 1
        Loop
        dec_offsets(i) = boost
        ' Making dynalloc more likely
        If boost > 0 Then
            dynalloc_logp = dynalloc_logp - 1
            If dynalloc_logp < 2 Then dynalloc_logp = 2     ' IMAX(2, dynalloc_logp-1)
        End If
    Next

    ' ALLOC(fine_quant, nbEBands, int);
    If tell + 6 * 8 <= total_bits Then
        alloc_trim = ec_dec_icdf(dec, trim_icdf, 0, 7)
    Else
        alloc_trim = 5
    End If

    bits = (len_ * 8) * 8 - ec_tell_frac(dec) - 1       ' (((opus_int32)len*8)<<BITRES) - ec_tell_frac(dec) - 1
    If isTransient <> 0 And LM >= 2 And bits >= (LM + 2) * 8 Then anti_collapse_rsv = 8 Else anti_collapse_rsv = 0
    bits = bits - anti_collapse_rsv

    ' ALLOC(pulses, nbEBands, int); ALLOC(fine_priority, nbEBands, int);

    codedBands = clt_compute_allocation(mode, start, end_, dec_offsets, dec_cap, _
          alloc_trim, intensity, dual_stereo, bits, balance, dec_pulses, _
          dec_fine_quant, dec_fine_priority, C, LM, dec, 0, 0, 0)

    unquant_fine_energy mode, start, end_, st.oldBandE, dec_fine_quant, dec, C

    c_ = 0
    Do
        k = c_ * (DECODE_BUFFER_SIZE + overlap)
        CopyMemory st.decode_mem(k), st.decode_mem(k + N), (DECODE_BUFFER_SIZE - N + overlap) * 4  ' OPUS_MOVE
        c_ = c_ + 1
    Loop While c_ < CC

    ' Decode fixed codebook
    ' ALLOC(collapse_masks, C*nbEBands, unsigned char);

    ' ALLOC(X, C*N, celt_norm);   /**< Interleaved normalised MDCTs */  (VB) celt_norm_buf from 0
    If C = 2 Then Y = N Else Y = -1
    quant_all_bands 0, mode, start, end_, 0, Y, dec_collapse_masks, _
          dec_pulses, shortBlocks, spread_decision, dual_stereo, intensity, dec_tf_res, _
          len_ * (8 * 8) - anti_collapse_rsv, balance, dec, LM, codedBands, st.rng, 0, _
          st.arch, st.disable_inv

    If anti_collapse_rsv > 0 Then
        anti_collapse_on = ec_dec_bits(dec, 1)
    End If
    ' (VB) the counters
    celt_vbCounts(0) = celt_vbCounts(0) + 1
    If isTransient <> 0 Then celt_vbCounts(1) = celt_vbCounts(1) + 1
    If anti_collapse_on <> 0 Then celt_vbCounts(2) = celt_vbCounts(2) + 1
    If C = 2 And intensity < end_ Then celt_vbCounts(3) = celt_vbCounts(3) + 1
    If dual_stereo <> 0 Then celt_vbCounts(4) = celt_vbCounts(4) + 1
    If silence <> 0 Then celt_vbCounts(5) = celt_vbCounts(5) + 1
    If st.downsample > 1 Then celt_vbCounts(11) = celt_vbCounts(11) + 1
    For i = start To end_ - 1
        If dec_offsets(i) > 0 Then celt_vbCounts(13) = celt_vbCounts(13) + 1: Exit For
    Next
    If codedBands < end_ Then celt_vbCounts(14) = celt_vbCounts(14) + 1

    unquant_energy_finalise mode, start, end_, st.oldBandE, _
          dec_fine_quant, dec_fine_priority, len_ * 8 - ec_tell(dec), dec, C

    If anti_collapse_on <> 0 Then
        anti_collapse mode, 0, dec_collapse_masks, LM, C, N, _
              start, end_, st.oldBandE, st.oldLogE, st.oldLogE2, dec_pulses, st.rng, st.arch
    End If

    If silence <> 0 Then
        For i = 0 To C * nbEBands - 1
            st.oldBandE(i) = -28!                       ' -QCONST16(28.f,DB_SHIFT)
        Next
    End If
    If st.prefilter_and_fold <> 0 Then
        prefilter_and_fold st, N
    End If
    celt_synthesis mode, 0, st.decode_mem, os0, os1, st.oldBandE, start, effEnd, _
                   C, CC, isTransient, LM, st.downsample, silence, st.arch

    c_ = 0
    Do
        If c_ = 0 Then os = os0 Else os = os1
        If st.postfilter_period < COMBFILTER_MINPERIOD Then st.postfilter_period = COMBFILTER_MINPERIOD
        If st.postfilter_period_old < COMBFILTER_MINPERIOD Then st.postfilter_period_old = COMBFILTER_MINPERIOD
        comb_filter st.decode_mem, os, st.decode_mem, os, st.postfilter_period_old, st.postfilter_period, celt_modes(mode).shortMdctSize, _
              st.postfilter_gain_old, st.postfilter_gain, st.postfilter_tapset_old, st.postfilter_tapset, _
              celt_modes(mode).window, overlap, st.arch
        If LM <> 0 Then
            comb_filter st.decode_mem, os + celt_modes(mode).shortMdctSize, st.decode_mem, os + celt_modes(mode).shortMdctSize, _
                  st.postfilter_period, postfilter_pitch, N - celt_modes(mode).shortMdctSize, _
                  st.postfilter_gain, postfilter_gain, st.postfilter_tapset, postfilter_tapset, _
                  celt_modes(mode).window, overlap, st.arch
        End If

        c_ = c_ + 1
    Loop While c_ < CC
    st.postfilter_period_old = st.postfilter_period
    st.postfilter_gain_old = st.postfilter_gain
    st.postfilter_tapset_old = st.postfilter_tapset
    st.postfilter_period = postfilter_pitch
    st.postfilter_gain = postfilter_gain
    st.postfilter_tapset = postfilter_tapset
    If LM <> 0 Then
        st.postfilter_period_old = st.postfilter_period
        st.postfilter_gain_old = st.postfilter_gain
        st.postfilter_tapset_old = st.postfilter_tapset
    End If

    If C = 1 Then
        CopyMemory st.oldBandE(nbEBands), st.oldBandE(0), nbEBands * 4    ' OPUS_COPY(&oldBandE[nbEBands], oldBandE, nbEBands);
    End If

    If isTransient = 0 Then
        CopyMemory st.oldLogE2(0), st.oldLogE(0), 2 * nbEBands * 4      ' OPUS_COPY(oldLogE2, oldLogE, 2*nbEBands);
        CopyMemory st.oldLogE(0), st.oldBandE(0), 2 * nbEBands * 4      ' OPUS_COPY(oldLogE, oldBandE, 2*nbEBands);
    Else
        For i = 0 To 2 * nbEBands - 1
            If st.oldBandE(i) < st.oldLogE(i) Then st.oldLogE(i) = st.oldBandE(i)  ' MIN16(oldLogE[i], oldBandE[i])
        Next
    End If
    ' In normal circumstances, we only allow the noise floor to increase by
    '   up to 2.4 dB/second, but when we're in DTX we give the weight of
    '   all missing packets to the update packet.
    k = st.loss_duration + M_
    If k > 160 Then k = 160
    max_background_increase = k * 0.001!               ' IMIN(160, st->loss_duration+M)*QCONST16(0.001f,DB_SHIFT)
    For i = 0 To 2 * nbEBands - 1
        ' MIN16(backgroundLogE[i] + max_background_increase, oldBandE[i])
        If st.backgroundLogE(i) + max_background_increase < st.oldBandE(i) Then
            st.backgroundLogE(i) = st.backgroundLogE(i) + max_background_increase
        Else
            st.backgroundLogE(i) = st.oldBandE(i)
        End If
    Next
    ' In case start or end were to change
    c_ = 0
    Do
        For i = 0 To start - 1
            st.oldBandE(c_ * nbEBands + i) = 0
            st.oldLogE(c_ * nbEBands + i) = -28!: st.oldLogE2(c_ * nbEBands + i) = -28!
        Next
        For i = end_ To nbEBands - 1
            st.oldBandE(c_ * nbEBands + i) = 0
            st.oldLogE(c_ * nbEBands + i) = -28!: st.oldLogE2(c_ * nbEBands + i) = -28!
        Next
        c_ = c_ + 1
    Loop While c_ < 2
    st.rng = dec.rng

    deemphasis st.decode_mem, os0, os1, pcm, pcmOffset, N, CC, st.downsample, celt_modes(mode).preemph, st.preemph_memD, accum
    st.loss_duration = 0
    st.prefilter_and_fold = 0
    If ec_tell(dec) > 8 * len_ Then
        decode_body_ = OPUS_INTERNAL_ERROR
        Exit Function
    End If
    If ec_get_error(dec) <> 0 Then st.error_ = 1
    decode_body_ = frame_size \ st.downsample
End Function

' (VB) data NULL <=> dataIsNull; ec_dec *dec NULL <=> decIsNull; pcm: pcm() from pcmOffset
Public Function celt_decode_with_ec(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, _
        ByVal dataIsNull As Boolean, ByVal len_ As Long, pcm() As Single, ByVal pcmOffset As Long, _
        ByVal frame_size As Long, dec As ec_ctx, ByVal decIsNull As Boolean, ByVal accum As Long) As Long
    celt_decode_with_ec = celt_decode_with_ec_dred(st, data, dataOffset, dataIsNull, len_, pcm, pcmOffset, frame_size, dec, decIsNull, accum)
    ' #ifdef ENABLE_DEEP_PLC , NULL #endif
End Function

' #ifdef CUSTOM_MODES: opus_custom_decode, opus_custom_decode_float (not compiled)

' opus_custom_decoder_ctl (celt.h: #define celt_decoder_ctl opus_custom_decoder_ctl)
' (VB) the va_arg is value: a GET writes it, a SET reads it (the C's NULL checks have no VB counterpart)
Public Function celt_decoder_ctl(ByVal st As Long, ByVal request As Long, Optional ByRef value As Long) As Long
    Dim i As Long, nbEBands As Long, ret As Long
    ret = OPUS_OK
    ' (VB) bad_arg: ret = OPUS_BAD_ARG; bad_request: ret = OPUS_UNIMPLEMENTED (no GoTo out of the With)
    With celt_decs(st)
        Select Case request
        Case OPUS_SET_COMPLEXITY_REQUEST
            If value < 0 Or value > 10 Then
                ret = OPUS_BAD_ARG
            Else
                .complexity = value
            End If
        Case OPUS_GET_COMPLEXITY_REQUEST
            value = .complexity
        Case CELT_SET_START_BAND_REQUEST
            If value < 0 Or value >= celt_modes(.mode).nbEBands Then
                ret = OPUS_BAD_ARG
            Else
                .start = value
            End If
        Case CELT_SET_END_BAND_REQUEST
            If value < 1 Or value > celt_modes(.mode).nbEBands Then
                ret = OPUS_BAD_ARG
            Else
                .end_ = value
            End If
        Case CELT_SET_CHANNELS_REQUEST
            If value < 1 Or value > 2 Then
                ret = OPUS_BAD_ARG
            Else
                .stream_channels = value
            End If
        Case CELT_GET_AND_CLEAR_ERROR_REQUEST
            value = .error_
            .error_ = 0
        Case OPUS_GET_LOOKAHEAD_REQUEST
            value = .overlap \ .downsample
        Case OPUS_RESET_STATE
            nbEBands = celt_modes(.mode).nbEBands
            ' OPUS_CLEAR((char*)&st->DECODER_RESET_START, opus_custom_decoder_get_size(st->mode, st->channels)-
            '       ((char*)&st->DECODER_RESET_START - (char*)st));
            .rng = 0
            .error_ = 0
            .last_pitch_index = 0
            .loss_duration = 0
            .skip_plc = 0
            .postfilter_period = 0
            .postfilter_period_old = 0
            .postfilter_gain = 0
            .postfilter_gain_old = 0
            .postfilter_tapset = 0
            .postfilter_tapset_old = 0
            .prefilter_and_fold = 0
            .preemph_memD(0) = 0
            .preemph_memD(1) = 0
            ZeroMemory .decode_mem(0), (UBound(.decode_mem) + 1) * 4
            ZeroMemory .lpc(0), (UBound(.lpc) + 1) * 4
            ZeroMemory .oldBandE(0), (UBound(.oldBandE) + 1) * 4
            ZeroMemory .backgroundLogE(0), (UBound(.backgroundLogE) + 1) * 4
            For i = 0 To 2 * nbEBands - 1
                .oldLogE(i) = -28!: .oldLogE2(i) = -28!     ' -QCONST16(28.f,DB_SHIFT)
            Next
            .skip_plc = 1
        Case OPUS_GET_PITCH_REQUEST
            value = .postfilter_period
        Case CELT_GET_MODE_REQUEST
            value = .mode
        Case CELT_SET_SIGNALLING_REQUEST
            .signalling = value
        Case OPUS_GET_FINAL_RANGE_REQUEST
            value = .rng
        Case OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST
            If value < 0 Or value > 1 Then
                ret = OPUS_BAD_ARG
            Else
                .disable_inv = value
            End If
        Case OPUS_GET_PHASE_INVERSION_DISABLED_REQUEST
            value = .disable_inv
        Case Else
            ret = OPUS_UNIMPLEMENTED                    ' goto bad_request
        End Select
    End With
    celt_decoder_ctl = ret
End Function
