// statsref_physfs.cpp: PhysFS for tools/statsref, over real folders: the
// search path (given by the command file, as Warzone.exe --vbtest prints
// its own), and a write folder (the writedir command: the replays), which
// as in Warzone comes first in the search path. Enough for WzConfig,
// loadFile, strres, modelGet and the replay files.
#include "physfs.h"

#include <windows.h>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <string>
#include <vector>

static std::vector<std::string> searchPath;
static std::string writeDir;
static PHYSFS_ErrorCode lastError = PHYSFS_ERR_OK;

void statsref_set_searchpath(const std::string &list)
{
	searchPath.clear();
	std::string cur;
	for (char c : list)
	{
		if (c == '|')
		{
			searchPath.push_back(cur);
			cur.clear();
		}
		else
		{
			cur += c;
		}
	}
	if (!cur.empty())
	{
		searchPath.push_back(cur);
	}
}

static std::wstring widen(const std::string &s);

// the write folder, put first in the search path
void statsref_set_writedir(const std::string &dir)
{
	writeDir = dir;
	CreateDirectoryW(widen(dir).c_str(), nullptr);
	searchPath.insert(searchPath.begin(), dir);
}

static std::wstring widen(const std::string &s)
{
	int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
	std::wstring w(n > 0 ? n - 1 : 0, L'\0');
	if (n > 1)
	{
		MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, &w[0], n);
	}
	return w;
}

static std::string nativePath(const std::string &dir, const char *name)
{
	std::string p = dir;
	if (!p.empty() && p.back() != '\\' && p.back() != '/')
	{
		p += '\\';
	}
	for (const char *c = name; *c; ++c)
	{
		p += (*c == '/') ? '\\' : *c;
	}
	return p;
}

// the first search path entry that has name: its index, the attributes; -1 if none
static int findEntry(const char *name, DWORD *attrs, std::string *full)
{
	for (size_t i = 0; i < searchPath.size(); i++)
	{
		std::string p = nativePath(searchPath[i], name);
		DWORD a = GetFileAttributesW(widen(p).c_str());
		if (a != INVALID_FILE_ATTRIBUTES)
		{
			if (attrs)
			{
				*attrs = a;
			}
			if (full)
			{
				*full = p;
			}
			return (int)i;
		}
	}
	lastError = PHYSFS_ERR_NOT_FOUND;
	return -1;
}

struct ShimFile
{
	FILE *f;
	PHYSFS_sint64 size;
};

extern "C" {

void PHYSFS_getLinkedVersion(PHYSFS_Version *ver)
{
	ver->major = PHYSFS_VER_MAJOR;
	ver->minor = PHYSFS_VER_MINOR;
	ver->patch = PHYSFS_VER_PATCH;
}

PHYSFS_File *PHYSFS_openRead(const char *filename)
{
	DWORD attrs;
	std::string full;
	if (findEntry(filename, &attrs, &full) < 0)
	{
		return nullptr;
	}
	if (attrs & FILE_ATTRIBUTE_DIRECTORY)
	{
		lastError = PHYSFS_ERR_NOT_A_FILE;
		return nullptr;
	}
	FILE *f = _wfopen(widen(full).c_str(), L"rb");
	if (!f)
	{
		lastError = PHYSFS_ERR_IO;
		return nullptr;
	}
	ShimFile *sf = new ShimFile;
	sf->f = f;
	_fseeki64(f, 0, SEEK_END);
	sf->size = _ftelli64(f);
	_fseeki64(f, 0, SEEK_SET);
	PHYSFS_File *h = new PHYSFS_File;
	h->opaque = sf;
	return h;
}

PHYSFS_File *PHYSFS_openWrite(const char *filename)
{
	if (writeDir.empty())
	{
		lastError = PHYSFS_ERR_NO_WRITE_DIR;
		return nullptr;
	}
	std::string full = nativePath(writeDir, filename);
	// the folders on the way
	for (size_t i = writeDir.size() + 1; i < full.size(); i++)
	{
		if (full[i] == '\\')
		{
			CreateDirectoryW(widen(full.substr(0, i)).c_str(), nullptr);
		}
	}
	FILE *f = _wfopen(widen(full).c_str(), L"wb");
	if (!f)
	{
		lastError = PHYSFS_ERR_IO;
		return nullptr;
	}
	ShimFile *sf = new ShimFile;
	sf->f = f;
	sf->size = 0;
	PHYSFS_File *h = new PHYSFS_File;
	h->opaque = sf;
	return h;
}

PHYSFS_File *PHYSFS_openAppend(const char *filename)
{
	lastError = PHYSFS_ERR_NO_WRITE_DIR;
	return nullptr;
}

int PHYSFS_close(PHYSFS_File *handle)
{
	ShimFile *sf = (ShimFile *)handle->opaque;
	fclose(sf->f);
	delete sf;
	delete handle;
	return 1;
}

PHYSFS_sint64 PHYSFS_readBytes(PHYSFS_File *handle, void *buffer, PHYSFS_uint64 len)
{
	ShimFile *sf = (ShimFile *)handle->opaque;
	return (PHYSFS_sint64)fread(buffer, 1, (size_t)len, sf->f);
}

PHYSFS_sint64 PHYSFS_writeBytes(PHYSFS_File *handle, const void *buffer, PHYSFS_uint64 len)
{
	ShimFile *sf = (ShimFile *)handle->opaque;
	return (PHYSFS_sint64)fwrite(buffer, 1, (size_t)len, sf->f);
}

PHYSFS_sint64 PHYSFS_read(PHYSFS_File *handle, void *buffer, PHYSFS_uint32 objSize, PHYSFS_uint32 objCount)
{
	PHYSFS_sint64 n = PHYSFS_readBytes(handle, buffer, (PHYSFS_uint64)objSize * objCount);
	return objSize ? n / objSize : 0;
}

PHYSFS_sint64 PHYSFS_write(PHYSFS_File *handle, const void *buffer, PHYSFS_uint32 objSize, PHYSFS_uint32 objCount)
{
	return -1;
}

PHYSFS_sint64 PHYSFS_fileLength(PHYSFS_File *handle)
{
	return ((ShimFile *)handle->opaque)->size;
}

PHYSFS_sint64 PHYSFS_tell(PHYSFS_File *handle)
{
	return _ftelli64(((ShimFile *)handle->opaque)->f);
}

int PHYSFS_seek(PHYSFS_File *handle, PHYSFS_uint64 pos)
{
	return _fseeki64(((ShimFile *)handle->opaque)->f, (long long)pos, SEEK_SET) == 0;
}

int PHYSFS_eof(PHYSFS_File *handle)
{
	ShimFile *sf = (ShimFile *)handle->opaque;
	return _ftelli64(sf->f) >= sf->size;
}

int PHYSFS_flush(PHYSFS_File *handle)
{
	return 1;
}

int PHYSFS_setBuffer(PHYSFS_File *handle, PHYSFS_uint64 bufsize)
{
	return 1;
}

int PHYSFS_exists(const char *fname)
{
	if (!*fname)
	{
		return 1;
	}
	return findEntry(fname, nullptr, nullptr) >= 0;
}

int PHYSFS_isDirectory(const char *fname)
{
	DWORD attrs;
	if (!*fname)
	{
		return 1;
	}
	return findEntry(fname, &attrs, nullptr) >= 0 && (attrs & FILE_ATTRIBUTE_DIRECTORY);
}

int PHYSFS_stat(const char *fname, PHYSFS_Stat *stat)
{
	DWORD attrs;
	std::string full;
	stat->filesize = -1;
	stat->modtime = -1;
	stat->createtime = -1;
	stat->accesstime = -1;
	stat->filetype = PHYSFS_FILETYPE_OTHER;
	stat->readonly = 1;
	if (!*fname)
	{
		stat->filesize = 0;
		stat->filetype = PHYSFS_FILETYPE_DIRECTORY;
		return 1;
	}
	if (findEntry(fname, &attrs, &full) < 0)
	{
		return 0;
	}
	if (attrs & FILE_ATTRIBUTE_DIRECTORY)
	{
		stat->filesize = 0;
		stat->filetype = PHYSFS_FILETYPE_DIRECTORY;
	}
	else
	{
		WIN32_FILE_ATTRIBUTE_DATA d;
		GetFileAttributesExW(widen(full).c_str(), GetFileExInfoStandard, &d);
		stat->filesize = ((PHYSFS_sint64)d.nFileSizeHigh << 32) | d.nFileSizeLow;
		stat->filetype = PHYSFS_FILETYPE_REGULAR;
	}
	return 1;
}

PHYSFS_sint64 PHYSFS_getLastModTime(const char *filename)
{
	return -1;
}

const char *PHYSFS_getRealDir(const char *filename)
{
	int i = findEntry(filename, nullptr, nullptr);
	return i < 0 ? nullptr : searchPath[i].c_str();
}

const char *PHYSFS_getWriteDir(void)
{
	return writeDir.empty() ? nullptr : writeDir.c_str();
}

int PHYSFS_setWriteDir(const char *newDir)
{
	return 1;
}

int PHYSFS_mount(const char *newDir, const char *mountPoint, int appendToPath)
{
	return 0;
}

const char *PHYSFS_getMountPoint(const char *dir)
{
	return "/";
}

char **PHYSFS_getSearchPath(void)
{
	char **list = (char **)calloc(searchPath.size() + 1, sizeof(char *));
	for (size_t i = 0; i < searchPath.size(); i++)
	{
		list[i] = strdup(searchPath[i].c_str());
	}
	return list;
}

const char *PHYSFS_getBaseDir(void)
{
	return ".\\";
}

const char *PHYSFS_getPrefDir(const char *org, const char *app)
{
	return nullptr;
}

const char *PHYSFS_getDirSeparator(void)
{
	return "\\";
}

int PHYSFS_mkdir(const char *dirName)
{
	return 0;
}

int PHYSFS_delete(const char *filename)
{
	return 0;
}

int PHYSFS_unmount(const char *oldDir)
{
	return 0;
}

int PHYSFS_removeFromSearchPath(const char *oldDir)
{
	return 0;
}

int PHYSFS_enumerate(const char *dir, PHYSFS_EnumerateCallback c, void *d)
{
	std::vector<std::string> seen;
	for (const std::string &root : searchPath)
	{
		std::string p = nativePath(root, dir);
		WIN32_FIND_DATAW fd;
		HANDLE h = FindFirstFileW(widen(p + "\\*").c_str(), &fd);
		if (h == INVALID_HANDLE_VALUE)
		{
			continue;
		}
		do
		{
			char name[MAX_PATH * 3];
			WideCharToMultiByte(CP_UTF8, 0, fd.cFileName, -1, name, sizeof(name), nullptr, nullptr);
			if (!strcmp(name, ".") || !strcmp(name, ".."))
			{
				continue;
			}
			bool dup = false;
			for (const std::string &s : seen)
			{
				dup = dup || s == name;
			}
			if (dup)
			{
				continue;
			}
			seen.push_back(name);
			PHYSFS_EnumerateCallbackResult r = c(d, dir, name);
			if (r != PHYSFS_ENUM_OK)
			{
				FindClose(h);
				return r == PHYSFS_ENUM_STOP;
			}
		}
		while (FindNextFileW(h, &fd));
		FindClose(h);
	}
	return 1;
}

static PHYSFS_EnumerateCallbackResult collect(void *d, const char *origdir, const char *fname)
{
	((std::vector<std::string> *)d)->push_back(fname);
	return PHYSFS_ENUM_OK;
}

char **PHYSFS_enumerateFiles(const char *dir)
{
	std::vector<std::string> names;
	PHYSFS_enumerate(dir, collect, &names);
	char **list = (char **)calloc(names.size() + 1, sizeof(char *));
	for (size_t i = 0; i < names.size(); i++)
	{
		list[i] = strdup(names[i].c_str());
	}
	return list;
}

void PHYSFS_freeList(void *listVar)
{
	char **list = (char **)listVar;
	if (!list)
	{
		return;
	}
	for (char **p = list; *p; ++p)
	{
		free(*p);
	}
	free(list);
}

PHYSFS_ErrorCode PHYSFS_getLastErrorCode(void)
{
	PHYSFS_ErrorCode e = lastError;
	lastError = PHYSFS_ERR_OK;
	return e;
}

const char *PHYSFS_getErrorByCode(PHYSFS_ErrorCode code)
{
	switch (code)
	{
	case PHYSFS_ERR_OK: return "no error";
	case PHYSFS_ERR_NOT_FOUND: return "not found";
	case PHYSFS_ERR_NOT_A_FILE: return "not a file";
	case PHYSFS_ERR_NO_WRITE_DIR: return "write directory is not set";
	case PHYSFS_ERR_IO: return "i/o error";
	default: return "unknown error";
	}
}

const char *PHYSFS_getLastError(void)
{
	return PHYSFS_getErrorByCode(PHYSFS_getLastErrorCode());
}

#define READ_INT(name, type) \
	int name(PHYSFS_File *file, type *val) \
	{ \
		type v; \
		if (PHYSFS_readBytes(file, &v, sizeof(v)) != sizeof(v)) return 0; \
		*val = v; \
		return 1; \
	}
READ_INT(PHYSFS_readSLE16, PHYSFS_sint16)
READ_INT(PHYSFS_readULE16, PHYSFS_uint16)
READ_INT(PHYSFS_readSLE32, PHYSFS_sint32)
READ_INT(PHYSFS_readULE32, PHYSFS_uint32)
#undef READ_INT

int PHYSFS_readSBE16(PHYSFS_File *file, PHYSFS_sint16 *val)
{
	PHYSFS_uint16 v;
	if (!PHYSFS_readULE16(file, &v)) return 0;
	*val = (PHYSFS_sint16)__builtin_bswap16(v);
	return 1;
}

int PHYSFS_readUBE16(PHYSFS_File *file, PHYSFS_uint16 *val)
{
	PHYSFS_uint16 v;
	if (!PHYSFS_readULE16(file, &v)) return 0;
	*val = __builtin_bswap16(v);
	return 1;
}

int PHYSFS_readSBE32(PHYSFS_File *file, PHYSFS_sint32 *val)
{
	PHYSFS_uint32 v;
	if (!PHYSFS_readULE32(file, &v)) return 0;
	*val = (PHYSFS_sint32)__builtin_bswap32(v);
	return 1;
}

int PHYSFS_readUBE32(PHYSFS_File *file, PHYSFS_uint32 *val)
{
	PHYSFS_uint32 v;
	if (!PHYSFS_readULE32(file, &v)) return 0;
	*val = __builtin_bswap32(v);
	return 1;
}

int PHYSFS_writeSLE16(PHYSFS_File *file, PHYSFS_sint16 val) { return 0; }
int PHYSFS_writeULE16(PHYSFS_File *file, PHYSFS_uint16 val) { return 0; }
int PHYSFS_writeSBE16(PHYSFS_File *file, PHYSFS_sint16 val) { return 0; }
int PHYSFS_writeUBE16(PHYSFS_File *file, PHYSFS_uint16 val) { return 0; }
int PHYSFS_writeSLE32(PHYSFS_File *file, PHYSFS_sint32 val) { return 0; }
int PHYSFS_writeULE32(PHYSFS_File *file, PHYSFS_uint32 val) { return 0; }
// (the replay files)
int PHYSFS_writeSBE32(PHYSFS_File *file, PHYSFS_sint32 val)
{
	return PHYSFS_writeUBE32(file, (PHYSFS_uint32)val);
}
int PHYSFS_writeUBE32(PHYSFS_File *file, PHYSFS_uint32 val)
{
	unsigned char b[4] = {(unsigned char)(val >> 24), (unsigned char)(val >> 16), (unsigned char)(val >> 8), (unsigned char)val};
	return PHYSFS_writeBytes(file, b, 4) == 4;
}

}

// lib/framework/wzapp.h's wzGetLogicalCPUCount (SDL_GetCPUCount: the logical CPUs), here for windows.h
uint32_t wzGetLogicalCPUCount()
{
	SYSTEM_INFO si;
	GetSystemInfo(&si);
	return si.dwNumberOfProcessors;
}
