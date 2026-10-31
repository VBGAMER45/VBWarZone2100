Attribute VB_Name = "theora_huffdec_c"
' ref/theora (libtheora 1.2.0) lib/huffdec.c with huffdec.h and huffman.h. The tables OC_DCT_TOKEN_MAP and
' OC_DCT_TOKEN_MAP_LOG_NENTRIES are in the generated theora_tables_vb.
' (VB) A set of trees (the C's ogg_int16_t *_nodes[TH_NHUFFMAN_TABLES], each tree malloc'd) is an array of
' offsets (nodes) into one Long array holding the trees one after another (store, nstore of it used); a
' tree's child indices stay relative to its own start, as in the C. A token list (unsigned char
' _tokens[256][2]) is a flat Long array, tokens(ti * 2) the token and tokens(ti * 2 + 1) its codeword
' length, passed with the index of its first entry where the C passes _tokens+ti.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' huffman.h
' The range of valid quantized DCT coefficient values. VP3 used 511 in the encoder, but the bitstream is
' capable of 580.
Public Const OC_DCT_VAL_RANGE As Long = 580

Public Const OC_NDCT_TOKEN_BITS As Long = 5

Public Const OC_DCT_EOB1_TOKEN As Long = 0
Public Const OC_DCT_EOB2_TOKEN As Long = 1
Public Const OC_DCT_EOB3_TOKEN As Long = 2
Public Const OC_DCT_REPEAT_RUN0_TOKEN As Long = 3
Public Const OC_DCT_REPEAT_RUN1_TOKEN As Long = 4
Public Const OC_DCT_REPEAT_RUN2_TOKEN As Long = 5
Public Const OC_DCT_REPEAT_RUN3_TOKEN As Long = 6

Public Const OC_DCT_SHORT_ZRL_TOKEN As Long = 7
Public Const OC_DCT_ZRL_TOKEN As Long = 8

Public Const OC_ONE_TOKEN As Long = 9
Public Const OC_MINUS_ONE_TOKEN As Long = 10
Public Const OC_TWO_TOKEN As Long = 11
Public Const OC_MINUS_TWO_TOKEN As Long = 12

Public Const OC_DCT_VAL_CAT2 As Long = 13
Public Const OC_DCT_VAL_CAT3 As Long = 17
Public Const OC_DCT_VAL_CAT4 As Long = 18
Public Const OC_DCT_VAL_CAT5 As Long = 19
Public Const OC_DCT_VAL_CAT6 As Long = 20
Public Const OC_DCT_VAL_CAT7 As Long = 21
Public Const OC_DCT_VAL_CAT8 As Long = 22

Public Const OC_DCT_RUN_CAT1A As Long = 23
Public Const OC_DCT_RUN_CAT1B As Long = 28
Public Const OC_DCT_RUN_CAT1C As Long = 29
Public Const OC_DCT_RUN_CAT2A As Long = 30
Public Const OC_DCT_RUN_CAT2B As Long = 31

Public Const OC_NDCT_EOB_TOKEN_MAX As Long = 7
Public Const OC_NDCT_ZRL_TOKEN_MAX As Long = 9
Public Const OC_NDCT_VAL_MAX As Long = 23
Public Const OC_NDCT_VAL_CAT1_MAX As Long = 13
Public Const OC_NDCT_VAL_CAT2_MAX As Long = 17
Public Const OC_NDCT_VAL_CAT2_SIZE As Long = OC_NDCT_VAL_CAT2_MAX - OC_DCT_VAL_CAT2
Public Const OC_NDCT_RUN_MAX As Long = 32
Public Const OC_NDCT_RUN_CAT1A_MAX As Long = 28

' Instead of storing every branching in the tree, subtrees can be collapsed into one node, with a table of
' size 1<<nbits pointing directly to its descedents nbits levels down. This allows more than one bit to be
' read at a time, and avoids following all the intermediate branches with next to no increased code
' complexity once the collapsed tree has been built. We do _not_ require that a subtree be complete to be
' collapsed, but instead store duplicate pointers in the table, and record the actual depth of the node
' below its parent. This tells us the number of bits to advance the stream after reaching it.
' This turns out to be equivalent to the method described in \cite{Hash95}, without the requirement that
' codewords be sorted by length. If the codewords were sorted by length (so-called ``canonical-codes''),
' they could be decoded much faster via either Lindell and Moffat's approach or Hashemian's Condensed
' Huffman Code approach, the latter of which has an extremely small memory footprint. We can't use Choueka
' et al.'s finite state machine approach, which is extremely fast, because we can't allow multiple symbols to
' be output at a time; the codebook can and does change between symbols. It also has very large memory
' requirements, which impairs cache coherency.
' We store the tree packed in an array of 16-bit integers (words). Each node consists of a single word,
' followed consecutively by two or more indices of its children. Let n be the value of this first word.
' This is the number of bits that need to be read to traverse the node, and must be positive. 1<<n entries
' follow in the array, each an index to a child node. If the child is positive, then it is the index of
' another internal node in the table. If the child is negative or zero, then it is a leaf node. These are
' stored directly in the child pointer to save space, since they only require a single word. If a leaf node
' would have been encountered before reading n bits, then it is duplicated the necessary number of times in
' this table. Leaf nodes pack both a token value and their actual depth in the tree. The token in the leaf
' node is (-leaf&255). The number of bits that need to be consumed to reach the leaf, starting from the
' current node, is (-leaf>>8).
'   @ARTICLE{Hash95,
'     author="Reza Hashemian",
'     title="Memory Efficient and High-Speed Search {Huffman} Coding",
'     journal="{IEEE} Transactions on Communications",
'     volume=43,
'     number=10,
'     pages="2576--2581",
'     month=Oct,
'     year=1995
'   }

' The map from external spec-defined tokens to internal tokens (OC_DCT_TOKEN_MAP) and the log base 2 of
' number of internal tokens associated with each of the spec tokens (OC_DCT_TOKEN_MAP_LOG_NENTRIES): see
' theora_tables_vb.

' The size a lookup table is allowed to grow to relative to the number of unique nodes it contains. E.g., if
' OC_HUFF_SLUSH is 4, then at most 75% of the space in the tree is wasted (1/4 of the space must be used).
' Larger numbers can decode tokens with fewer read operations, while smaller numbers may save more space.
' With a sample file:
' 32233473 read calls are required when no tree collapsing is done (100.0%).
' 19269269 read calls are required when OC_HUFF_SLUSH is 1 (59.8%).
' 11144969 read calls are required when OC_HUFF_SLUSH is 2 (34.6%).
' 10538563 read calls are required when OC_HUFF_SLUSH is 4 (32.7%).
' 10192578 read calls are required when OC_HUFF_SLUSH is 8 (31.6%).
' Since a value of 2 gets us the vast majority of the speed-up with only a small amount of wasted memory,
' this is what we use. This value must be less than 128, or you could create a tree with more than 32767
' entries, which would overflow the 16-bit words used to index it.
Private Const OC_HUFF_SLUSH As Long = 2
' The root of the tree is on the fast path, and a larger value here is more beneficial than elsewhere in the
' tree. 7 appears to give the best performance, trading off between increased use of the single-read fast
' path and cache footprint for the tables, though obviously this will depend on your cache size. Using 7
' here, the VP3 tables are about twice as large compared to using 2.
Private Const OC_ROOT_HUFF_SLUSH As Long = 7

' (VB) 0x80000000U>>_n for n = 0..31, as a Double (the C adds them up in an ogg_uint32_t)
Private Function oc_vbTopBit(ByVal n As Long) As Double
    If n = 0 Then oc_vbTopBit = 2147483648# Else oc_vbTopBit = OC_PB_POW2(31 - n)
End Function

' Unpacks a Huffman codebook.
'  _opb:    The buffer to unpack from.
'  _tokens: Stores a list of internal tokens, in the order they were found in the codebook, and the lengths
'            of their corresponding codewords. This is enough to completely define the codebook, while
'            minimizing stack usage and avoiding temporary allocations (for platforms where free() is a
'            no-op).
'  Return: The number of internal tokens in the codebook, or a negative value on error.
Public Function oc_huff_tree_unpack(opb As oc_pack_buf, tokens() As Long) As Long
    Dim code As Long        ' ogg_uint32_t
    Dim len_ As Long
    Dim ntokens As Long
    Dim nleaves As Long
    Dim bits As Long
    Dim code_bit As Long    ' ogg_uint32_t
    Dim neb As Long
    Dim nentries As Long
    Dim token As Long
    code = 0
    len_ = 0
    ntokens = 0
    nleaves = 0
    Do
        bits = oc_pack_read1(opb)
        ' Only process nodes so long as there's more bits in the buffer.
        If oc_pack_bytes_left(opb) < 0 Then oc_huff_tree_unpack = TH_EBADHEADER: Exit Function
        ' Read an internal node:
        If bits = 0 Then
            len_ = len_ + 1
            ' Don't allow codewords longer than 32 bits.
            If len_ > 32 Then oc_huff_tree_unpack = TH_EBADHEADER: Exit Function
        ' Read a leaf node:
        Else
            ' Don't allow more than 32 spec-tokens per codebook.
            nleaves = nleaves + 1
            If nleaves > 32 Then oc_huff_tree_unpack = TH_EBADHEADER: Exit Function
            bits = oc_pack_read(opb, OC_NDCT_TOKEN_BITS)
            neb = OC_DCT_TOKEN_MAP_LOG_NENTRIES(bits)
            token = OC_DCT_TOKEN_MAP(bits)
            nentries = OC_PB_POW2(neb)
            Do While nentries > 0
                nentries = nentries - 1
                tokens(ntokens * 2) = token And 255
                token = token + 1
                tokens(ntokens * 2 + 1) = (len_ + neb) And 255
                ntokens = ntokens + 1
            Loop
            If len_ <= 0 Then Exit Do
            ' code_bit=0x80000000U>>len-1
            If len_ = 1 Then code_bit = &H80000000 Else code_bit = OC_PB_POW2(32 - len_)
            Do While len_ > 0
                If (code And code_bit) = 0 Then Exit Do
                code = code Xor code_bit
                ' code_bit<<=1 (unsigned: the top bit shifts out)
                If code_bit = &H80000000 Then
                    code_bit = 0
                ElseIf code_bit = &H40000000 Then
                    code_bit = &H80000000
                Else
                    code_bit = code_bit * 2
                End If
                len_ = len_ - 1
            Loop
            If len_ <= 0 Then Exit Do
            code = code Or code_bit
        End If
    Loop
    oc_huff_tree_unpack = ntokens
End Function

' Count how many tokens would be required to fill a subtree at depth _depth.
'  _tokens: A list of internal tokens, in the order they are found in the codebook, and the lengths of their
'            corresponding codewords.
'  _depth:  The depth of the desired node in the corresponding tree structure.
'  Return: The number of tokens that belong to that subtree.
Private Function oc_huff_subtree_tokens(tokens() As Long, ByVal tokOffset As Long, ByVal depth As Long) As Long
    Dim code As Double      ' ogg_uint32_t: (VB) a Double, below 2^32
    Dim ti As Long
    code = 0
    ti = 0
    Do
        If tokens((tokOffset + ti) * 2 + 1) - depth < 32 Then
            code = code + oc_vbTopBit(tokens((tokOffset + ti) * 2 + 1) - depth)
            ti = ti + 1
        Else
            ' Because of the expanded internal tokens, we can have codewords as long as 35 bits. A single
            ' recursion here is enough to advance past them.
            code = code + 1
            ti = ti + oc_huff_subtree_tokens(tokens, tokOffset + ti, depth + 31)
        End If
    Loop While code < 2147483648#
    oc_huff_subtree_tokens = ti
End Function

' Compute the number of bits to use for a collapsed tree node at the given depth.
'  _tokens:  A list of internal tokens, in the order they are found in the codebook, and the lengths of their
'             corresponding codewords.
'  _ntokens: The number of tokens corresponding to this tree node.
'  _depth:   The depth of this tree node.
'  Return: The number of bits to use for a collapsed tree node rooted here. This is always at least one, even
'   if this was a leaf node.
Private Function oc_huff_tree_collapse_depth(tokens() As Long, ByVal tokOffset As Long, ByVal ntokens As Long, _
    ByVal depth As Long) As Long
    Dim got_leaves As Long
    Dim loccupancy As Long
    Dim occupancy As Long
    Dim slush As Long
    Dim nbits As Long
    Dim best_nbits As Long
    Dim ti As Long
    If depth > 0 Then slush = OC_HUFF_SLUSH Else slush = OC_ROOT_HUFF_SLUSH
    ' It's legal to have a tree with just a single node, which requires no bits to decode and always returns
    ' the same token. However, no encoder actually does this (yet). To avoid a special case in
    ' oc_huff_token_decode(), we force the number of lookahead bits to be at least one. This will produce a
    ' tree that looks ahead one bit and then advances the stream zero bits.
    nbits = 1
    occupancy = 2
    got_leaves = 1
    Do
        If got_leaves <> 0 Then best_nbits = nbits
        nbits = nbits + 1
        got_leaves = 0
        loccupancy = occupancy
        occupancy = 0
        ti = 0
        Do While ti < ntokens
            If tokens((tokOffset + ti) * 2 + 1) < depth + nbits Then
                ti = ti + 1
            ElseIf tokens((tokOffset + ti) * 2 + 1) = depth + nbits Then
                got_leaves = 1
                ti = ti + 1
            Else
                ti = ti + oc_huff_subtree_tokens(tokens, tokOffset + ti, depth + nbits)
            End If
            occupancy = occupancy + 1
        Loop
        If occupancy <= loccupancy Then Exit Do
    Loop While occupancy * slush >= OC_PB_POW2(nbits)
    oc_huff_tree_collapse_depth = best_nbits
End Function

' Determines the size in words of a Huffman tree node that represents a subtree of depth _nbits.
'  _nbits: The depth of the subtree. This must be greater than zero.
'  Return: The number of words required to store the node.
Private Function oc_huff_node_size(ByVal nbits As Long) As Long
    oc_huff_node_size = 1 + OC_PB_POW2(nbits)
End Function

' Produces a collapsed-tree representation of the given token list.
'  _tree: The storage for the collapsed Huffman tree. This may be NULL to compute the required storage size
'          instead of constructing the tree. (VB) tree(treeOffset), treeIsNull for NULL
'  _tokens:  A list of internal tokens, in the order they are found in the codebook, and the lengths of their
'             corresponding codewords.
'  _ntokens: The number of tokens corresponding to this tree node.
'  Return: The number of words required to store the tree.
Private Function oc_huff_tree_collapse(tree() As Long, ByVal treeOffset As Long, ByVal treeIsNull As Boolean, _
    tokens() As Long, ByVal ntokens As Long) As Long
    Dim node(0 To 33) As Long       ' ogg_int16_t
    Dim depth(0 To 33) As Long      ' unsigned char
    Dim last(0 To 33) As Long       ' unsigned char
    Dim ntree As Long
    Dim ti As Long
    Dim l As Long
    Dim nbits As Long
    Dim leaf As Long
    Dim nentries As Long
    depth(0) = 0
    last(0) = (ntokens - 1) And 255
    ntree = 0
    ti = 0
    l = 0
    Do
        nbits = oc_huff_tree_collapse_depth(tokens, ti, last(l) + 1 - ti, depth(l))
        node(l) = ntree
        ntree = ntree + oc_huff_node_size(nbits)
        If Not treeIsNull Then tree(treeOffset + node(l)) = nbits: node(l) = node(l) + 1
        Do
            Do While ti <= last(l)
                If tokens(ti * 2 + 1) > depth(l) + nbits Then Exit Do
                If Not treeIsNull Then
                    nentries = OC_PB_POW2(depth(l) + nbits - tokens(ti * 2 + 1))
                    leaf = -((tokens(ti * 2 + 1) - depth(l)) * 256 Or tokens(ti * 2))
                    Do While nentries > 0
                        nentries = nentries - 1
                        tree(treeOffset + node(l)) = leaf
                        node(l) = node(l) + 1
                    Loop
                End If
                ti = ti + 1
            Loop
            If ti <= last(l) Then
                ' We need to recurse
                depth(l + 1) = (depth(l) + nbits) And 255
                If Not treeIsNull Then tree(treeOffset + node(l)) = ntree: node(l) = node(l) + 1
                l = l + 1
                last(l) = (ti + oc_huff_subtree_tokens(tokens, ti, depth(l)) - 1) And 255
                Exit Do
            ' Pop back up a level of recursion.
            Else
                l = l - 1                   ' else if(l-->0)nbits=depth[l+1]-depth[l];
                If l + 1 > 0 Then nbits = depth(l + 1) - depth(l)
            End If
        Loop While l >= 0
    Loop While l >= 0
    oc_huff_tree_collapse = ntree
End Function

' Unpacks a set of Huffman trees, and reduces them to a collapsed representation.
'  _opb:   The buffer to unpack the trees from.
'  _nodes: The table to fill with the Huffman trees. (VB) nodes(i) the offset of tree i in store
'  Return: 0 on success, or a negative value on error. The caller is responsible for cleaning up any
'           partially initialized _nodes on failure.
Public Function oc_huff_trees_unpack(opb As oc_pack_buf, nodes() As Long, store() As Long, nstore As Long) As Long
    Dim i As Long
    Dim tokens(0 To 511) As Long    ' unsigned char tokens[256][2]
    Dim ntokens As Long
    Dim size As Long
    Dim dummy() As Long
    For i = 0 To TH_NHUFFMAN_TABLES - 1
        ' Unpack the full tree into a temporary buffer.
        ntokens = oc_huff_tree_unpack(opb, tokens)
        If ntokens < 0 Then oc_huff_trees_unpack = ntokens: Exit Function
        ' Figure out how big the collapsed tree will be and allocate space for it.
        size = oc_huff_tree_collapse(dummy, 0, True, tokens, ntokens)
        ' This should never happen; if it does it means you set OC_HUFF_SLUSH or OC_ROOT_HUFF_SLUSH too large.
        If size > 32767 Then oc_huff_trees_unpack = TH_EIMPL: Exit Function
        ' tree=(ogg_int16_t *)_ogg_malloc(size*sizeof(*tree)): (VB) the next size words of store
        If nstore + size > oc_vbStoreSize(store) Then ReDim Preserve store(0 To 2 * (nstore + size) - 1)
        ' Construct the collapsed the tree.
        oc_huff_tree_collapse store, nstore, False, tokens, ntokens
        nodes(i) = nstore
        nstore = nstore + size
    Next
    oc_huff_trees_unpack = 0
End Function

' (VB) the number of entries of a tree store (0 when it was never allocated)
Private Function oc_vbStoreSize(store() As Long) As Long
    On Error GoTo none
    oc_vbStoreSize = UBound(store) + 1
    Exit Function
none:
    oc_vbStoreSize = 0
End Function

' Determines the size in words of a Huffman subtree.
'  _tree: The complete Huffman tree.
'  _node: The index of the root of the desired subtree.
'  Return: The number of words required to store the tree.
Private Function oc_huff_tree_size(tree() As Long, ByVal treeOffset As Long, ByVal node As Long) As Long
    Dim size As Long
    Dim nchildren As Long
    Dim n As Long
    Dim i As Long
    Dim child As Long
    n = tree(treeOffset + node)
    size = oc_huff_node_size(n)
    nchildren = OC_PB_POW2(n)
    i = 0
    Do
        child = tree(treeOffset + node + i + 1)
        If child <= 0 Then
            i = i + OC_PB_POW2(n - (-child \ 256))
        Else
            size = size + oc_huff_tree_size(tree, treeOffset, child)
            i = i + 1
        End If
    Loop While i < nchildren
    oc_huff_tree_size = size
End Function

' Makes a copy of the given set of Huffman trees.
'  _dst: The array to store the copy in.
'  _src: The array of trees to copy.
' (VB) dst trees into dstStore (nDstStore used, from 0), src trees in srcStore
Public Function oc_huff_trees_copy(dst() As Long, dstStore() As Long, nDstStore As Long, src() As Long, _
    srcStore() As Long) As Long
    Dim i As Long
    Dim size As Long
    Dim total As Long
    total = 0
    For i = 0 To TH_NHUFFMAN_TABLES - 1
        total = total + oc_huff_tree_size(srcStore, src(i), 0)
    Next
    ReDim dstStore(0 To total - 1)
    nDstStore = 0
    For i = 0 To TH_NHUFFMAN_TABLES - 1
        size = oc_huff_tree_size(srcStore, src(i), 0)
        dst(i) = nDstStore
        CopyMemory dstStore(nDstStore), srcStore(src(i)), size * 4
        nDstStore = nDstStore + size
    Next
    oc_huff_trees_copy = 0
End Function

' Frees the memory used by a set of Huffman trees.
'  _nodes: The array of trees to free.
Public Sub oc_huff_trees_clear(nodes() As Long, store() As Long, nstore As Long)
    Dim i As Long
    For i = 0 To TH_NHUFFMAN_TABLES - 1
        nodes(i) = 0
    Next
    Erase store
    nstore = 0
End Sub

' Unpacks a single token using the given Huffman tree.
'  _opb:  The buffer to unpack the token from.
'  _node: The tree to unpack the token with. (VB) tree(treeOffset)
'  Return: The token value.
' (oc_huff_token_decode: bitpack.h maps it to oc_huff_token_decode_c)
Public Function oc_huff_token_decode(opb As oc_pack_buf, tree() As Long, ByVal treeOffset As Long) As Long
    Dim ptr As Long
    Dim stop_ As Long
    Dim window As Long
    Dim available As Long
    Dim bits As Long
    Dim node As Long
    Dim n As Long
    Dim shift As Long
    ptr = opb.ptr
    window = opb.window
    stop_ = opb.stop_
    available = opb.bits
    node = 0
    Do
        n = tree(treeOffset + node)
        If n > available Then
            shift = OC_PB_WINDOW_SIZE - available
            Do
                ' We don't bother setting eof because we won't check for it after we've started decoding DCT
                ' tokens.
                If ptr >= stop_ Then
                    shift = -OC_LOTS_OF_BITS
                    Exit Do
                End If
                shift = shift - 8
                window = window Or (opb.buf(ptr) * OC_PB_POW2(shift))
                ptr = ptr + 1
            Loop While shift >= 8
            ' Note: We never request more than 24 bits, so there's no need to fill in the last partial byte
            ' here.
            available = OC_PB_WINDOW_SIZE - shift
        End If
        bits = window \ OC_PB_POW2(OC_PB_WINDOW_SIZE - n)
        node = tree(treeOffset + node + 1 + bits)
        If node <= 0 Then Exit Do
        window = (window And OC_PB_MASK(OC_PB_WINDOW_SIZE - n)) * OC_PB_POW2(n)
        available = available - n
    Loop
    node = -node
    n = node \ 256
    window = (window And OC_PB_MASK(OC_PB_WINDOW_SIZE - n)) * OC_PB_POW2(n)
    available = available - n
    opb.ptr = ptr
    opb.window = window
    opb.bits = available
    oc_huff_token_decode = node And 255
End Function
