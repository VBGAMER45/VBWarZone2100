// mapscriptref: reference output of wzmaplib's script-generated maps (Warzone 4.7.0's lib/wzmaplib with
// map_script.cpp and the QuickJS it pins, ref/quickjs-wz at a57c627) for tools/map_script_test.py, which builds
// it with g++ / gcc and compares its output with the port's (MapScriptTests.exe over WzMap.dll).
//
//   mapscriptref <command file> <output folder>
//
// Each line of the command file: "folder|maxplayers|seed" <TAB> output (relative to the output folder): the
// folder loaded by Map::loadFromPath as a skirmish map with that seed (a uint32_t), then a text dump of the map
// data, the objects, the CRCs and every log line the library printed (tools/mapref.cpp's dump, the script's
// log() lines among them).
#include "../ref/warzone2100/lib/wzmaplib/include/wzmaplib/map.h"
#include <cstdarg>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <sstream>
#include <string>

static std::string out;

class CaptureLogger : public WzMap::LoggingProtocol
{
public:
	void printLog(LogLevel level, const char *function, int line, const char *str) override
	{
		char buf[64];
		snprintf(buf, sizeof(buf), "log %d %s:%d ", static_cast<int>(level), function, line);
		out += buf;
		out += str;
		if (out.empty() || out.back() != '\n') out += "\n";
	}
};

static void put(const char *fmt, ...) __attribute__((format(printf, 1, 2)));
static void put(const char *fmt, ...)
{
	char buf[4096];
	va_list ap;
	va_start(ap, fmt);
	vsnprintf(buf, sizeof(buf), fmt, ap);
	va_end(ap);
	out += buf;
}

static std::vector<std::string> split(const std::string& s, char c)
{
	std::vector<std::string> parts;
	std::stringstream ss(s);
	std::string p;
	while (std::getline(ss, p, c)) parts.push_back(p);
	return parts;
}

static void dumpMap(const std::string& arg)
{
	auto parts = split(arg, '|');
	auto logger = std::make_shared<CaptureLogger>();
	uint32_t seed = static_cast<uint32_t>(std::stoul(parts[2]));
	auto map = WzMap::Map::loadFromPath(parts[0], WzMap::MapType::SKIRMISH, std::stoul(parts[1]), seed, logger);
	if (!map)
	{
		put("loadFromPath failed\n");
		return;
	}
	uint32_t gotSeed = 0;
	put("scriptgenerated %d seed %u\n", map->wasScriptGenerated() ? 1 : 0, map->scriptGeneratedMapSeed().value_or(gotSeed));
	auto data = map->mapData();
	if (data)
	{
		put("mapdata %u %u tiles %zu gateways %zu\n", data->width, data->height, data->mMapTiles.size(), data->mGateways.size());
		for (size_t i = 0; i < data->mMapTiles.size(); i++)
		{
			put("%u:%u%s", data->mMapTiles[i].height, data->mMapTiles[i].texture, (i % 16 == 15 || i + 1 == data->mMapTiles.size()) ? "\n" : " ");
		}
	}
	else
	{
		put("mapdata none\n");
	}
	auto structures = map->mapStructures();
	if (structures)
	{
		put("structures %zu\n", structures->size());
		for (const auto& s : *structures)
		{
			put("s %s %u %u %u %s %d %u\n", s.name.c_str(), s.position.x, s.position.y, s.direction,
				s.id.has_value() ? std::to_string(s.id.value()).c_str() : "-", s.player, s.modules);
		}
	}
	else
	{
		put("structures none\n");
	}
	auto droids = map->mapDroids();
	if (droids)
	{
		put("droids %zu\n", droids->size());
		for (const auto& d : *droids)
		{
			put("d %s %u %u %u %s %d\n", d.name.c_str(), d.position.x, d.position.y, d.direction,
				d.id.has_value() ? std::to_string(d.id.value()).c_str() : "-", d.player);
		}
	}
	else
	{
		put("droids none\n");
	}
	auto features = map->mapFeatures();
	if (features)
	{
		put("features %zu\n", features->size());
		for (const auto& f : *features)
		{
			put("f %s %u %u %u %s %s\n", f.name.c_str(), f.position.x, f.position.y, f.direction,
				f.id.has_value() ? std::to_string(f.id.value()).c_str() : "-",
				f.player.has_value() ? std::to_string(f.player.value()).c_str() : "-");
		}
	}
	else
	{
		put("features none\n");
	}
	uint32_t crcTiles = map->crcSumMapTiles(0);
	uint32_t crcStructures = map->crcSumStructures(0);
	uint32_t crcDroids = map->crcSumDroids(0);
	uint32_t crcFeatures = map->crcSumFeatures(0);
	put("crc %08x %08x %08x %08x\n", crcTiles, crcStructures, crcDroids, crcFeatures);
}

int main(int argc, char** argv)
{
	if (argc < 3) return 2;
	std::ifstream cmds(argv[1]);
	std::string line, root = argv[2];
	while (std::getline(cmds, line))
	{
		if (!line.empty() && line.back() == '\r') line.pop_back();
		auto parts = split(line, '\t');
		if (parts.size() != 2) continue;
		out.clear();
		dumpMap(parts[0]);
		std::ofstream f(root + "/" + parts[1], std::ios::binary);
		f << out;
	}
	return 0;
}
