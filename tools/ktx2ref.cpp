// ktx2ref - the KTX2 reference (cluster T): Basis Universal's transcoder as 4.7.0 builds it
// (ref/basis_universal @ df079ec, with its zstddeclib.c), compiled by g++.
//   ktx2ref <format> <out dir> <file.ktx2>...
// format: a basist::transcoder_texture_format number (6 = BC7, 10 = ASTC 4x4, 13 = RGBA32, 3 = BC3,
// 2 = BC1, 4 = BC4, 5 = BC5, ...). For each file, every level (layer 0, face 0) is transcoded as
// gfx_api_image_basis_priv.cpp does (decode_flags 0; RGBA32 at the level's original size, block formats as
// m_total_blocks blocks) and written one after another to <out dir>/<file name>.<format>; a line per file
// gives its size, levels, the output's FNV-1a hash and the time in ms ("-" as out dir: hashes only).
#include <basisu_transcoder.h>
#include <chrono>
#include <cstdio>
#include <cstring>
#include <string>
#include <vector>

static bool readFile(const char *path, std::vector<uint8_t> &data)
{
	FILE *f = fopen(path, "rb");
	if (!f) return false;
	fseek(f, 0, SEEK_END);
	long n = ftell(f);
	fseek(f, 0, SEEK_SET);
	data.resize(n);
	bool ok = fread(data.data(), 1, n, f) == (size_t)n;
	fclose(f);
	return ok;
}

int main(int argc, char **argv)
{
	if (argc < 4)
	{
		fprintf(stderr, "usage: ktx2ref <format> <out dir> <file.ktx2>...\n");
		return 2;
	}
	basist::basisu_transcoder_init();
	auto format = (basist::transcoder_texture_format)atoi(argv[1]);
	std::string outDir = argv[2];
	int failures = 0;
	for (int a = 3; a < argc; a++)
	{
		std::vector<uint8_t> data;
		std::string path = argv[a];
		std::string name = path.substr(path.find_last_of("/\\") + 1);
		if (!readFile(argv[a], data))
		{
			printf("%s: cannot read\n", name.c_str());
			failures++;
			continue;
		}
		auto t0 = std::chrono::steady_clock::now();
		basist::ktx2_transcoder transcoder;
		std::vector<uint8_t> out;
		bool ok = transcoder.init(data.data(), (uint32_t)data.size()) && transcoder.start_transcoding();
		uint32_t levels = ok ? transcoder.get_levels() : 0;
		for (uint32_t level = 0; ok && level < levels; level++)
		{
			basist::ktx2_image_level_info info;
			if (!transcoder.get_image_level_info(info, level, 0, 0)) { ok = false; break; }
			uint32_t n = basist::basis_transcoder_format_is_uncompressed(format) ? info.m_orig_width * info.m_orig_height : info.m_total_blocks;
			size_t at = out.size();
			out.resize(at + (size_t)n * basist::basis_get_bytes_per_block_or_pixel(format));
			ok = transcoder.transcode_image_level(level, 0, 0, out.data() + at, n, format, 0);
		}
		double ms = std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - t0).count();
		if (!ok)
		{
			printf("%s: transcode failed\n", name.c_str());
			failures++;
			continue;
		}
		uint32_t h = 2166136261u;
		for (uint8_t b : out) h = (h ^ b) * 16777619u;
		printf("%s %ux%u levels %u bytes %zu fnv %08x %.2f ms\n", name.c_str(), transcoder.get_width(), transcoder.get_height(), levels, out.size(), h, ms);
		if (outDir != "-")
		{
			std::string o = outDir + "/" + name + "." + argv[1];
			FILE *f = fopen(o.c_str(), "wb");
			if (!f || fwrite(out.data(), 1, out.size(), f) != out.size()) { printf("%s: cannot write\n", o.c_str()); failures++; }
			if (f) fclose(f);
		}
	}
	return failures ? 1 : 0;
}
