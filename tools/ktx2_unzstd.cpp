// ktx2_unzstd - rewrites Zstandard-supercompressed KTX2 files without supercompression (cluster T, fork T2's test
// of the KTX2 path before the port's zstd decoder exists): each level is decompressed with zstd 1.4.9's decoder
// (ref/basis_universal/zstd/zstddeclib.c) and stored as it is; the header's supercompression scheme becomes 0 (none),
// each level index entry gets the level's new offset and byte length = uncompressed byte length; the DFD and the
// key/value data stay where they are; the levels follow them in the file's own order, each at a 16-byte boundary.
// A file already without supercompression is copied as it is.
//   ktx2_unzstd <out dir> <file.ktx2>...
#include <zstd.h>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <string>
#include <vector>
#include <algorithm>

static uint32_t rd32(const uint8_t *p) { return p[0] | (p[1] << 8) | (p[2] << 16) | ((uint32_t)p[3] << 24); }
static uint64_t rd64(const uint8_t *p) { return rd32(p) | ((uint64_t)rd32(p + 4) << 32); }
static void wr64(uint8_t *p, uint64_t v) { for (int i = 0; i < 8; i++) p[i] = (uint8_t)(v >> (i * 8)); }
static void wr32(uint8_t *p, uint32_t v) { for (int i = 0; i < 4; i++) p[i] = (uint8_t)(v >> (i * 8)); }

int main(int argc, char **argv)
{
	if (argc < 3)
	{
		fprintf(stderr, "usage: ktx2_unzstd <out dir> <file.ktx2>...\n");
		return 2;
	}
	int failures = 0;
	for (int a = 2; a < argc; a++)
	{
		std::string path = argv[a];
		std::string name = path.substr(path.find_last_of("/\\") + 1);
		FILE *f = fopen(argv[a], "rb");
		if (!f) { printf("%s: cannot read\n", name.c_str()); failures++; continue; }
		fseek(f, 0, SEEK_END);
		long n = ftell(f);
		fseek(f, 0, SEEK_SET);
		std::vector<uint8_t> in(n);
		bool ok = fread(in.data(), 1, n, f) == (size_t)n;
		fclose(f);
		if (!ok || n < 80) { printf("%s: cannot read\n", name.c_str()); failures++; continue; }
		uint32_t levels = rd32(&in[40]);
		uint32_t scheme = rd32(&in[44]);
		std::vector<uint8_t> out;
		if (scheme == 0)
		{
			out = in;
		}
		else if (scheme == 2)
		{
			// everything before the first level's data, then the levels in file order
			uint64_t first = UINT64_MAX;
			std::vector<uint32_t> order(levels);
			for (uint32_t i = 0; i < levels; i++)
			{
				order[i] = i;
				first = std::min(first, rd64(&in[80 + 24 * i]));
			}
			std::sort(order.begin(), order.end(), [&](uint32_t x, uint32_t y) { return rd64(&in[80 + 24 * x]) < rd64(&in[80 + 24 * y]); });
			out.assign(in.begin(), in.begin() + (size_t)first);
			wr32(&out[44], 0);
			for (uint32_t i : order)
			{
				const uint8_t *e = &in[80 + 24 * i];
				uint64_t ofs = rd64(e), len = rd64(e + 8), ulen = rd64(e + 16);
				while (out.size() % 16) out.push_back(0);
				uint64_t at = out.size();
				out.resize(at + ulen);
				size_t r = ZSTD_decompress(out.data() + at, ulen, in.data() + ofs, len);
				if (ZSTD_isError(r) || r != ulen) { ok = false; break; }
				uint8_t *oe = &out[80 + 24 * i];
				wr64(oe, at);
				wr64(oe + 8, ulen);
				wr64(oe + 16, ulen);
			}
			if (!ok) { printf("%s: zstd failed\n", name.c_str()); failures++; continue; }
		}
		else
		{
			printf("%s: supercompression scheme %u not handled\n", name.c_str(), scheme);
			failures++;
			continue;
		}
		std::string o = std::string(argv[1]) + "/" + name;
		FILE *g = fopen(o.c_str(), "wb");
		if (!g || fwrite(out.data(), 1, out.size(), g) != out.size()) { printf("%s: cannot write\n", o.c_str()); failures++; }
		if (g) fclose(g);
	}
	return failures ? 1 : 0;
}
