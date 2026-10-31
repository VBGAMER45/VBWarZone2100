Attribute VB_Name = "silk_sort_c"
' libopus 1.5.2 (ref/opus/silk): sort.c, the decoder's function (silk_insertion_sort_increasing_all_values_int16,
' for NLSF_stabilize.c; silk_insertion_sort_increasing and silk_insertion_sort_decreasing_int16 are the
' encoder's).
Option Explicit

Public Sub silk_insertion_sort_increasing_all_values_int16(a() As Integer, ByVal L As Long)
    Dim value As Long
    Dim i As Long, j As Long

    '/* Safety checks */
    ' celt_assert( L >  0 );

    '/* Sort vector elements by value, increasing order */
    For i = 1 To L - 1
        value = a(i)
        j = i - 1
        Do While j >= 0
            If Not (value < a(j)) Then Exit Do
            a(j + 1) = a(j)         '/* Shift value */
            j = j - 1
        Loop
        a(j + 1) = value            '/* Write value */
    Next
End Sub
