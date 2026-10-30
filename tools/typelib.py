"""Read a COM type library (a VB6 ActiveX DLL) without the registry.

  python -I tools\\typelib.py <dll>     prints the library and its creatable classes

library(path) returns (name, tlbid, (major, minor), [(class name, clsid), ...])
for the creatable classes (coclasses with TYPEFLAG_FCANCREATE). The build uses
it to write the exe's Reference lines and its reg-free COM manifest.
"""
import ctypes
import sys
from ctypes import POINTER, WINFUNCTYPE, byref, c_int, c_long, c_uint, c_ushort, c_void_p

REGKIND_NONE = 2
TKIND_COCLASS = 5
TYPEFLAG_FCANCREATE = 0x2


class GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_uint32), ("Data2", ctypes.c_uint16), ("Data3", ctypes.c_uint16),
                ("Data4", ctypes.c_ubyte * 8)]

    def __str__(self):
        d = bytes(self.Data4)
        return "{%08X-%04X-%04X-%s-%s}" % (self.Data1, self.Data2, self.Data3, d[:2].hex().upper(), d[2:].hex().upper())


class TLIBATTR(ctypes.Structure):
    _fields_ = [("guid", GUID), ("lcid", ctypes.c_uint32), ("syskind", c_int),
                ("wMajorVerNum", c_ushort), ("wMinorVerNum", c_ushort), ("wLibFlags", c_ushort)]


class TYPEATTR(ctypes.Structure):       # leading fields only
    _fields_ = [("guid", GUID), ("lcid", ctypes.c_uint32), ("dwReserved", ctypes.c_uint32),
                ("memidConstructor", c_long), ("memidDestructor", c_long), ("lpstrSchema", c_void_p),
                ("cbSizeInstance", ctypes.c_uint32), ("typekind", c_int), ("cFuncs", c_ushort),
                ("cVars", c_ushort), ("cImplTypes", c_ushort), ("cbSizeVft", c_ushort),
                ("cbAlignment", c_ushort), ("wTypeFlags", c_ushort)]


oleaut32 = ctypes.WinDLL("oleaut32")
oleaut32.LoadTypeLibEx.argtypes = [ctypes.c_wchar_p, c_int, POINTER(c_void_p)]
oleaut32.LoadTypeLibEx.restype = c_long
oleaut32.SysFreeString.argtypes = [c_void_p]


def _method(obj, index, *argtypes):
    vtbl = ctypes.cast(obj, POINTER(POINTER(c_void_p)))[0]
    return WINFUNCTYPE(c_long, c_void_p, *argtypes)(vtbl[index])


def _bstr(p):
    s = ctypes.wstring_at(p.value) if p.value else ""
    oleaut32.SysFreeString(p)
    return s


def library(path):
    tl = c_void_p()
    hr = oleaut32.LoadTypeLibEx(path, REGKIND_NONE, byref(tl))
    if hr != 0:
        raise OSError("LoadTypeLibEx(%s) failed: 0x%08X" % (path, hr & 0xFFFFFFFF))
    try:
        pattr = POINTER(TLIBATTR)()
        _method(tl, 7, POINTER(POINTER(TLIBATTR)))(tl, byref(pattr))           # GetLibAttr
        attr = pattr.contents
        tlbid, version = str(attr.guid), (attr.wMajorVerNum, attr.wMinorVerNum)
        _method(tl, 12, POINTER(TLIBATTR))(tl, pattr)                           # ReleaseTLibAttr
        get_doc = _method(tl, 9, c_int, POINTER(c_void_p), c_void_p, c_void_p, c_void_p)
        name = c_void_p()
        get_doc(tl, -1, byref(name), None, None, None)
        libname = _bstr(name)
        count = WINFUNCTYPE(c_uint, c_void_p)(ctypes.cast(tl, POINTER(POINTER(c_void_p)))[0][3])(tl)
        classes = []
        for i in range(count):
            kind = c_int()
            _method(tl, 5, c_uint, POINTER(c_int))(tl, i, byref(kind))         # GetTypeInfoType
            if kind.value != TKIND_COCLASS:
                continue
            ti = c_void_p()
            _method(tl, 4, c_uint, POINTER(c_void_p))(tl, i, byref(ti))        # GetTypeInfo
            pta = POINTER(TYPEATTR)()
            _method(ti, 3, POINTER(POINTER(TYPEATTR)))(ti, byref(pta))          # GetTypeAttr
            clsid, flags = str(pta.contents.guid), pta.contents.wTypeFlags
            _method(ti, 19, POINTER(TYPEATTR))(ti, pta)                         # ReleaseTypeAttr
            _method(ti, 2)(ti)                                                  # Release
            if flags & TYPEFLAG_FCANCREATE:
                cname = c_void_p()
                get_doc(tl, i, byref(cname), None, None, None)
                classes.append((_bstr(cname), clsid))
        return libname, tlbid, version, classes
    finally:
        _method(tl, 2)(tl)                                                      # Release


if __name__ == "__main__":
    libname, tlbid, version, classes = library(sys.argv[1])
    print("%s %s %d.%d" % (libname, tlbid, version[0], version[1]))
    for cname, clsid in classes:
        print("  %s.%s %s" % (libname, cname, clsid))
