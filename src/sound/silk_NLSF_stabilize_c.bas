Attribute VB_Name = "silk_NLSF_stabilize_c"
' libopus 1.5.2 (ref/opus/silk): NLSF_stabilize.c. (VB) NDeltaMin_Q15 is an offset into silk_rom_i16; the C's
' I (beside i) is I_.
Option Explicit

'/* NLSF stabilizer:                                         */
'/*                                                          */
'/* - Moves NLSFs further apart if they are too close        */
'/* - Moves NLSFs away from borders if they are too close    */
'/* - High effort to achieve a modification with minimum     */
'/*     Euclidean distance to input vector                   */
'/* - Output are sorted NLSF coefficients                    */
'/*                                                          */

'/* Constant Definitions */
Private Const MAX_LOOPS As Long = 20

'/* NLSF stabilizer, for a single input data vector */
Public Sub silk_NLSF_stabilize(NLSF_Q15() As Integer, ByVal NDeltaMin_Q15 As Long, ByVal L As Long)
    Dim i As Long, I_ As Long, k As Long, loops As Long
    Dim center_freq_Q15 As Integer
    Dim diff_Q15 As Long, min_diff_Q15 As Long, min_center_Q15 As Long, max_center_Q15 As Long

    '/* This is necessary to ensure an output within range of a opus_int16 */
    ' silk_assert( NDeltaMin_Q15[L] >= 1 );

    For loops = 0 To MAX_LOOPS - 1
        '/**************************/
        '/* Find smallest distance */
        '/**************************/
        '/* First element */
        min_diff_Q15 = CLng(NLSF_Q15(0)) - silk_rom_i16(NDeltaMin_Q15)
        I_ = 0
        '/* Middle elements */
        For i = 1 To L - 1
            diff_Q15 = NLSF_Q15(i) - (CLng(NLSF_Q15(i - 1)) + silk_rom_i16(NDeltaMin_Q15 + i))
            If diff_Q15 < min_diff_Q15 Then
                min_diff_Q15 = diff_Q15
                I_ = i
            End If
        Next
        '/* Last element */
        diff_Q15 = &H8000& - (CLng(NLSF_Q15(L - 1)) + silk_rom_i16(NDeltaMin_Q15 + L))
        If diff_Q15 < min_diff_Q15 Then
            min_diff_Q15 = diff_Q15
            I_ = L
        End If

        '/***************************************************/
        '/* Now check if the smallest distance non-negative */
        '/***************************************************/
        If min_diff_Q15 >= 0 Then
            Exit Sub
        End If

        If I_ = 0 Then
            '/* Move away from lower limit */
            NLSF_Q15(0) = silk_rom_i16(NDeltaMin_Q15)

        ElseIf I_ = L Then
            '/* Move away from higher limit */
            NLSF_Q15(L - 1) = &H8000& - silk_rom_i16(NDeltaMin_Q15 + L)

        Else
            '/* Find the lower extreme for the location of the current center frequency */
            min_center_Q15 = 0
            For k = 0 To I_ - 1
                min_center_Q15 = min_center_Q15 + silk_rom_i16(NDeltaMin_Q15 + k)
            Next
            min_center_Q15 = min_center_Q15 + silk_RSHIFT(silk_rom_i16(NDeltaMin_Q15 + I_), 1)

            '/* Find the upper extreme for the location of the current center frequency */
            max_center_Q15 = &H8000&
            For k = L To I_ + 1 Step -1
                max_center_Q15 = max_center_Q15 - silk_rom_i16(NDeltaMin_Q15 + k)
            Next
            max_center_Q15 = max_center_Q15 - silk_RSHIFT(silk_rom_i16(NDeltaMin_Q15 + I_), 1)

            '/* Move apart, sorted by value, keeping the same center frequency */
            center_freq_Q15 = silk_LIMIT(silk_RSHIFT_ROUND(CLng(NLSF_Q15(I_ - 1)) + NLSF_Q15(I_), 1), _
                min_center_Q15, max_center_Q15)
            NLSF_Q15(I_ - 1) = center_freq_Q15 - silk_RSHIFT(silk_rom_i16(NDeltaMin_Q15 + I_), 1)
            NLSF_Q15(I_) = CLng(NLSF_Q15(I_ - 1)) + silk_rom_i16(NDeltaMin_Q15 + I_)
        End If
    Next

    '/* Safe and simple fall back method, which is less ideal than the above */
    If loops = MAX_LOOPS Then
        '/* Insertion sort (fast for already almost sorted arrays):   */
        '/* Best case:  O(n)   for an already sorted array            */
        '/* Worst case: O(n^2) for an inversely sorted array          */
        silk_insertion_sort_increasing_all_values_int16 NLSF_Q15, L

        '/* First NLSF should be no less than NDeltaMin[0] */
        NLSF_Q15(0) = silk_max_int(NLSF_Q15(0), silk_rom_i16(NDeltaMin_Q15))

        '/* Keep delta_min distance between the NLSFs */
        For i = 1 To L - 1
            NLSF_Q15(i) = silk_max_int(NLSF_Q15(i), silk_ADD_SAT16(NLSF_Q15(i - 1), silk_rom_i16(NDeltaMin_Q15 + i)))
        Next

        '/* Last NLSF should be no higher than 1 - NDeltaMin[L] */
        NLSF_Q15(L - 1) = silk_min_int(NLSF_Q15(L - 1), &H8000& - silk_rom_i16(NDeltaMin_Q15 + L))

        '/* Keep NDeltaMin distance between the NLSFs */
        For i = L - 2 To 0 Step -1
            NLSF_Q15(i) = silk_min_int(NLSF_Q15(i), CLng(NLSF_Q15(i + 1)) - silk_rom_i16(NDeltaMin_Q15 + i + 1))
        Next
    End If
End Sub
