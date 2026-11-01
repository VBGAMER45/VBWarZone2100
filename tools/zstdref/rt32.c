/* A freestanding runtime for the 32-bit reference DLL (no CRT: this gcc has no 32-bit libraries). The heap is a
   static arena: ZSTD_decompress allocates its one DCtx, uses it and frees it. */
#include <stddef.h>
void* memcpy(void* d, const void* s, size_t n) { volatile unsigned char* dp = d; const unsigned char* sp = s; while (n--) *dp++ = *sp++; return d; }
void* memmove(void* d, const void* s, size_t n) { volatile unsigned char* dp = d; const unsigned char* sp = s; if ((const unsigned char*)dp < sp) { while (n--) *dp++ = *sp++; } else { dp += n; sp += n; while (n--) *--dp = *--sp; } return d; }
void* memset(void* d, int c, size_t n) { volatile unsigned char* dp = d; while (n--) *dp++ = (unsigned char)c; return d; }
static unsigned char arena[1 << 20] __attribute__((aligned(16)));
static size_t used;
void* malloc(size_t n) { void* p; n = (n + 15) & ~(size_t)15; if (used + n > sizeof(arena)) return 0; p = arena + used; used += n; return p; }
void* calloc(size_t n, size_t m) { void* p = malloc(n * m); if (p) memset(p, 0, n * m); return p; }
void free(void* p) { if (p == arena) used = 0; }
int __stdcall DllMainCRTStartup(void* h, unsigned long r, void* p) { (void)h; (void)r; (void)p; return 1; }
/* libgcc's stack probe (cygwin.S, __chkstk_ms): touches each 4 KB page of a large frame, top down */
__asm__(".globl ___chkstk_ms\n"
        "___chkstk_ms:\n"
        "  pushl %ecx\n  pushl %eax\n  cmpl $0x1000, %eax\n  leal 12(%esp), %ecx\n  jb 2f\n"
        "1:\n  subl $0x1000, %ecx\n  orl $0x0, (%ecx)\n  subl $0x1000, %eax\n  cmpl $0x1000, %eax\n  ja 1b\n"
        "2:\n  subl %eax, %ecx\n  orl $0x0, (%ecx)\n  popl %eax\n  popl %ecx\n  ret\n");
