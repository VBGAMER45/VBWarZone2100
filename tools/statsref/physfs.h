// physfs.h for tools/statsref: the part of the PhysFS 3 API that Warzone's
// framework headers and the sources statsref compiles use, implemented in
// statsref_physfs.cpp over a list of real folders (the search path).
#ifndef STATSREF_PHYSFS_H
#define STATSREF_PHYSFS_H

#include <stddef.h>

typedef unsigned char PHYSFS_uint8;
typedef signed char PHYSFS_sint8;
typedef unsigned short PHYSFS_uint16;
typedef signed short PHYSFS_sint16;
typedef unsigned int PHYSFS_uint32;
typedef signed int PHYSFS_sint32;
typedef unsigned long long PHYSFS_uint64;
typedef signed long long PHYSFS_sint64;

#define PHYSFS_VER_MAJOR 3
#define PHYSFS_VER_MINOR 2
#define PHYSFS_VER_PATCH 0
#define PHYSFS_DEPRECATED

typedef struct PHYSFS_File
{
	void *opaque;
} PHYSFS_File;
#define PHYSFS_file PHYSFS_File

typedef struct PHYSFS_Version
{
	PHYSFS_uint8 major;
	PHYSFS_uint8 minor;
	PHYSFS_uint8 patch;
} PHYSFS_Version;

typedef enum PHYSFS_FileType
{
	PHYSFS_FILETYPE_REGULAR,
	PHYSFS_FILETYPE_DIRECTORY,
	PHYSFS_FILETYPE_SYMLINK,
	PHYSFS_FILETYPE_OTHER
} PHYSFS_FileType;

typedef struct PHYSFS_Stat
{
	PHYSFS_sint64 filesize;
	PHYSFS_sint64 modtime;
	PHYSFS_sint64 createtime;
	PHYSFS_sint64 accesstime;
	PHYSFS_FileType filetype;
	int readonly;
} PHYSFS_Stat;

typedef enum PHYSFS_ErrorCode
{
	PHYSFS_ERR_OK,
	PHYSFS_ERR_OTHER_ERROR,
	PHYSFS_ERR_OUT_OF_MEMORY,
	PHYSFS_ERR_NOT_INITIALIZED,
	PHYSFS_ERR_IS_INITIALIZED,
	PHYSFS_ERR_ARGV0_IS_NULL,
	PHYSFS_ERR_UNSUPPORTED,
	PHYSFS_ERR_PAST_EOF,
	PHYSFS_ERR_FILES_STILL_OPEN,
	PHYSFS_ERR_INVALID_ARGUMENT,
	PHYSFS_ERR_NOT_MOUNTED,
	PHYSFS_ERR_NOT_FOUND,
	PHYSFS_ERR_SYMLINK_FORBIDDEN,
	PHYSFS_ERR_NO_WRITE_DIR,
	PHYSFS_ERR_OPEN_FOR_READING,
	PHYSFS_ERR_OPEN_FOR_WRITING,
	PHYSFS_ERR_NOT_A_FILE,
	PHYSFS_ERR_READ_ONLY,
	PHYSFS_ERR_CORRUPT,
	PHYSFS_ERR_SYMLINK_LOOP,
	PHYSFS_ERR_IO,
	PHYSFS_ERR_PERMISSION,
	PHYSFS_ERR_NO_SPACE,
	PHYSFS_ERR_BAD_FILENAME,
	PHYSFS_ERR_BUSY,
	PHYSFS_ERR_DIR_NOT_EMPTY,
	PHYSFS_ERR_OS_ERROR,
	PHYSFS_ERR_DUPLICATE,
	PHYSFS_ERR_BAD_PASSWORD,
	PHYSFS_ERR_APP_CALLBACK
} PHYSFS_ErrorCode;

typedef enum PHYSFS_EnumerateCallbackResult
{
	PHYSFS_ENUM_ERROR = -1,
	PHYSFS_ENUM_STOP = 0,
	PHYSFS_ENUM_OK = 1
} PHYSFS_EnumerateCallbackResult;

typedef PHYSFS_EnumerateCallbackResult (*PHYSFS_EnumerateCallback)(void *data, const char *origdir, const char *fname);

#ifdef __cplusplus
extern "C" {
#endif

void PHYSFS_getLinkedVersion(PHYSFS_Version *ver);
PHYSFS_File *PHYSFS_openRead(const char *filename);
PHYSFS_File *PHYSFS_openWrite(const char *filename);
PHYSFS_File *PHYSFS_openAppend(const char *filename);
int PHYSFS_close(PHYSFS_File *handle);
PHYSFS_sint64 PHYSFS_readBytes(PHYSFS_File *handle, void *buffer, PHYSFS_uint64 len);
PHYSFS_sint64 PHYSFS_writeBytes(PHYSFS_File *handle, const void *buffer, PHYSFS_uint64 len);
PHYSFS_sint64 PHYSFS_read(PHYSFS_File *handle, void *buffer, PHYSFS_uint32 objSize, PHYSFS_uint32 objCount);
PHYSFS_sint64 PHYSFS_write(PHYSFS_File *handle, const void *buffer, PHYSFS_uint32 objSize, PHYSFS_uint32 objCount);
PHYSFS_sint64 PHYSFS_fileLength(PHYSFS_File *handle);
PHYSFS_sint64 PHYSFS_tell(PHYSFS_File *handle);
int PHYSFS_seek(PHYSFS_File *handle, PHYSFS_uint64 pos);
int PHYSFS_eof(PHYSFS_File *handle);
int PHYSFS_flush(PHYSFS_File *handle);
int PHYSFS_setBuffer(PHYSFS_File *handle, PHYSFS_uint64 bufsize);
int PHYSFS_exists(const char *fname);
int PHYSFS_isDirectory(const char *fname);
int PHYSFS_stat(const char *fname, PHYSFS_Stat *stat);
PHYSFS_sint64 PHYSFS_getLastModTime(const char *filename);
const char *PHYSFS_getRealDir(const char *filename);
const char *PHYSFS_getWriteDir(void);
int PHYSFS_setWriteDir(const char *newDir);
int PHYSFS_mount(const char *newDir, const char *mountPoint, int appendToPath);
const char *PHYSFS_getMountPoint(const char *dir);
char **PHYSFS_getSearchPath(void);
const char *PHYSFS_getBaseDir(void);
const char *PHYSFS_getPrefDir(const char *org, const char *app);
const char *PHYSFS_getDirSeparator(void);
int PHYSFS_mkdir(const char *dirName);
int PHYSFS_delete(const char *filename);
int PHYSFS_unmount(const char *oldDir);
int PHYSFS_removeFromSearchPath(const char *oldDir);
char **PHYSFS_enumerateFiles(const char *dir);
int PHYSFS_enumerate(const char *dir, PHYSFS_EnumerateCallback c, void *d);
void PHYSFS_freeList(void *listVar);
PHYSFS_ErrorCode PHYSFS_getLastErrorCode(void);
const char *PHYSFS_getErrorByCode(PHYSFS_ErrorCode code);
const char *PHYSFS_getLastError(void);
int PHYSFS_readSLE16(PHYSFS_File *file, PHYSFS_sint16 *val);
int PHYSFS_readULE16(PHYSFS_File *file, PHYSFS_uint16 *val);
int PHYSFS_readSBE16(PHYSFS_File *file, PHYSFS_sint16 *val);
int PHYSFS_readUBE16(PHYSFS_File *file, PHYSFS_uint16 *val);
int PHYSFS_readSLE32(PHYSFS_File *file, PHYSFS_sint32 *val);
int PHYSFS_readULE32(PHYSFS_File *file, PHYSFS_uint32 *val);
int PHYSFS_readSBE32(PHYSFS_File *file, PHYSFS_sint32 *val);
int PHYSFS_readUBE32(PHYSFS_File *file, PHYSFS_uint32 *val);
int PHYSFS_writeSLE16(PHYSFS_File *file, PHYSFS_sint16 val);
int PHYSFS_writeULE16(PHYSFS_File *file, PHYSFS_uint16 val);
int PHYSFS_writeSBE16(PHYSFS_File *file, PHYSFS_sint16 val);
int PHYSFS_writeUBE16(PHYSFS_File *file, PHYSFS_uint16 val);
int PHYSFS_writeSLE32(PHYSFS_File *file, PHYSFS_sint32 val);
int PHYSFS_writeULE32(PHYSFS_File *file, PHYSFS_uint32 val);
int PHYSFS_writeSBE32(PHYSFS_File *file, PHYSFS_sint32 val);
int PHYSFS_writeUBE32(PHYSFS_File *file, PHYSFS_uint32 val);

#ifdef __cplusplus
}
#endif

#endif
