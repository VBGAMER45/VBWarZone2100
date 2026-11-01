/* The 32-bit reference of fork T1's zstd test: zstddeclib.c (zstd 1.4.9's decoder as Basis Universal ships it)
   compiled with gcc -m32, so its code paths are the 32-bit build's (MEM_32bits). Exported stdcall for VB. */
#include "zstddeclib.c"
__declspec(dllexport) size_t __stdcall z32_decompress(void* dst, size_t dstCapacity, const void* src, size_t srcSize)
{ return ZSTD_decompress(dst, dstCapacity, src, srcSize); }
