// jsonref: reference output of nlohmann::json 3.11.3 (Warzone's copy, from
// ref/warzone2100/3rdparty/json) for the VB port's tests (tools/run_tests.py
// builds it with g++ and compares its output with WzTests.exe's).
//
//   jsonref <command file> <output folder>
//
// Each line of the command file: mode <TAB> input file <TAB> output file
// (relative to the output folder). Modes:
//   dump      json::parse, dump(4) + "\n" (as WzConfig writes a file)
//   dumpo     ordered_json::parse, dump(4) + "\n"
//   compact   json::parse, dump()
//   ascii     json::parse, dump(-1, ' ', true)
//   error     json::parse: "ok", or the exception's what()
//   doubles   each input line a double's 16 hex digit bit pattern: json(d).dump(), one per line
// A failed parse in the dump modes writes "error: " + what().
#include <nlohmann/json.hpp>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <sstream>
#include <string>
using json = nlohmann::json;

static std::string readAll(const std::string& p)
{
	std::ifstream f(p, std::ios::binary);
	std::stringstream s;
	s << f.rdbuf();
	return s.str();
}

static void writeAll(const std::string& p, const std::string& s)
{
	std::ofstream f(p, std::ios::binary);
	f << s;
}

static std::string run(const std::string& mode, const std::string& in)
{
	std::string data = readAll(in);
	if (mode == "doubles")
	{
		std::istringstream lines(data);
		std::string line, out;
		while (std::getline(lines, line))
		{
			if (!line.empty() && line.back() == '\r') line.pop_back();
			if (line.empty()) continue;
			uint64_t bits = std::stoull(line, nullptr, 16);
			double d;
			std::memcpy(&d, &bits, 8);
			out += json(d).dump() + "\n";
		}
		return out;
	}
	try
	{
		if (mode == "dumpo") return nlohmann::ordered_json::parse(data).dump(4) + "\n";
		json j = json::parse(data);
		if (mode == "dump") return j.dump(4) + "\n";
		if (mode == "compact") return j.dump();
		if (mode == "ascii") return j.dump(-1, ' ', true);
		if (mode == "error") return "ok\n";
	}
	catch (const std::exception& e)
	{
		if (mode == "error") return std::string(e.what()) + "\n";
		return std::string("error: ") + e.what() + "\n";
	}
	return "unknown mode " + mode + "\n";
}

int main(int argc, char** argv)
{
	if (argc < 3) return 2;
	std::ifstream cmds(argv[1]);
	std::string line, root = argv[2];
	while (std::getline(cmds, line))
	{
		if (!line.empty() && line.back() == '\r') line.pop_back();
		size_t t1 = line.find('\t'), t2 = line.find('\t', t1 + 1);
		if (t1 == std::string::npos || t2 == std::string::npos) continue;
		writeAll(root + "/" + line.substr(t2 + 1), run(line.substr(0, t1), line.substr(t1 + 1, t2 - t1 - 1)));
	}
	return 0;
}
