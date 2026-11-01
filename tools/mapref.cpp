// mapref: reference output of wzmaplib (Warzone 4.7.0's lib/wzmaplib, from
// ref/warzone2100) for the VB port's tests (tools/run_tests.py builds it with
// g++ and compares its output with WzTests.exe's).
//
//   mapref <command file> <output folder>
//
// Each line of the command file: mode <TAB> argument <TAB> output (relative
// to the output folder). Modes:
//   map      argument "folder|maptype|maxplayers": Map::loadFromPath, then a
//            text dump of the loaded format, map data, objects, terrain
//            types, the CRCs and every log line the library printed
//   export   argument "folder|maptype|maxplayers|format": the map loaded and
//            exported with Map::exportMapToPath into the output folder (a
//            folder); the runner compares the files written
// Script-generated maps run 4.7.0's map_script.cpp over the QuickJS it pins
// (ref/quickjs-wz, built by tools/run_tests.py).
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

static WzMap::MapType mapTypeOf(const std::string& s)
{
	if (s == "campaign") return WzMap::MapType::CAMPAIGN;
	if (s == "savegame") return WzMap::MapType::SAVEGAME;
	return WzMap::MapType::SKIRMISH;
}

static void dumpMap(const std::string& arg)
{
	auto parts = split(arg, '|');
	auto logger = std::make_shared<CaptureLogger>();
	auto map = WzMap::Map::loadFromPath(parts[0], mapTypeOf(parts[1]), std::stoul(parts[2]), 0, logger);
	if (!map)
	{
		put("loadFromPath failed\n");
		return;
	}
	auto format = map->loadedMapFormat();
	put("format %d\n", format.has_value() ? static_cast<int>(format.value()) : -1);
	auto data = map->mapData();
	if (data)
	{
		put("mapdata %u %u tiles %zu gateways %zu\n", data->width, data->height, data->mMapTiles.size(), data->mGateways.size());
		for (size_t i = 0; i < data->mMapTiles.size(); i++)
		{
			put("%u:%u%s", data->mMapTiles[i].height, data->mMapTiles[i].texture, (i % 16 == 15 || i + 1 == data->mMapTiles.size()) ? "\n" : " ");
		}
		for (const auto& g : data->mGateways)
		{
			put("gateway %u %u %u %u\n", g.x1, g.y1, g.x2, g.y2);
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
	auto ttypes = map->mapTerrainTypes();
	if (ttypes)
	{
		put("terraintypes %zu\n", ttypes->terrainTypes.size());
		for (size_t i = 0; i < ttypes->terrainTypes.size(); i++)
		{
			put("%d%s", static_cast<int>(ttypes->terrainTypes[i]), (i % 32 == 31 || i + 1 == ttypes->terrainTypes.size()) ? "\n" : " ");
		}
	}
	else
	{
		put("terraintypes none\n");
	}
	// one at a time: a part that failed to load is loaded again (and logs again), in this order
	uint32_t crcTiles = map->crcSumMapTiles(0);
	uint32_t crcStructures = map->crcSumStructures(0);
	uint32_t crcDroids = map->crcSumDroids(0);
	uint32_t crcFeatures = map->crcSumFeatures(0);
	put("crc %08x %08x %08x %08x\n", crcTiles, crcStructures, crcDroids, crcFeatures);
}

static void exportMap(const std::string& arg, const std::string& outFolder)
{
	auto parts = split(arg, '|');
	auto logger = std::make_shared<CaptureLogger>();
	auto mapType = mapTypeOf(parts[1]);
	uint32_t maxPlayers = std::stoul(parts[2]);
	auto map = WzMap::Map::loadFromPath(parts[0], mapType, maxPlayers, 0, logger);
	if (!map)
	{
		put("loadFromPath failed\n");
		return;
	}
	auto format = static_cast<WzMap::OutputFormat>(std::stoi(parts[3]));
	bool ok = WzMap::Map::exportMapToPath(*map, outFolder, mapType, maxPlayers, format, logger);
	put("export %s\n", ok ? "ok" : "failed");
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
		if (parts.size() != 3) continue;
		out.clear();
		if (parts[0] == "map")
		{
			dumpMap(parts[1]);
			std::ofstream f(root + "/" + parts[2], std::ios::binary);
			f << out;
		}
		else if (parts[0] == "export")
		{
			exportMap(parts[1], root + "/" + parts[2]);
			std::ofstream f(root + "/" + parts[2] + ".log", std::ios::binary);
			f << out;
		}
	}
	return 0;
}
