"""Pure-ctypes COM helpers for Windows (no third-party libraries)."""

import ctypes
from typing import Optional

_ole32 = ctypes.windll.ole32
_oleaut32 = ctypes.windll.oleaut32

PTR_SIZE = ctypes.sizeof(ctypes.c_void_p)


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def guid(s: str) -> GUID:
    """Parse a ``{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}`` string into a GUID."""
    s = s.strip("{}")
    p = s.split("-")
    d4 = p[3] + p[4]
    return GUID(
        int(p[0], 16),
        int(p[1], 16),
        int(p[2], 16),
        (ctypes.c_ubyte * 8)(*[int(d4[i : i + 2], 16) for i in range(0, 16, 2)]),
    )


_vc_cache: dict = {}


def vc(p: int, idx: int, restype, argtypes: tuple, *args):
    """Call COM vtable slot *idx* on raw pointer *p* (int)."""
    if not p:
        raise ValueError("NULL COM pointer")
    key = (restype, argtypes)
    proto = _vc_cache.get(key)
    if proto is None:
        proto = ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)
        _vc_cache[key] = proto
    vtbl = ctypes.c_void_p.from_address(p).value
    faddr = ctypes.c_void_p.from_address(vtbl + idx * PTR_SIZE).value
    return proto(faddr)(p, *args)


def release(p: int) -> None:
    """Call IUnknown::Release (vtable slot 2)."""
    if p:
        vc(p, 2, ctypes.c_ulong, ())


def co_create_instance(clsid: GUID, iid: GUID) -> int:
    """CoCreateInstance with CLSCTX_INPROC_SERVER; returns raw pointer."""
    _ole32.CoInitialize(None)
    out = ctypes.c_void_p()
    hr = _ole32.CoCreateInstance(
        ctypes.byref(clsid),
        None,
        1,  # CLSCTX_INPROC_SERVER
        ctypes.byref(iid),
        ctypes.byref(out),
    )
    if hr < 0 or not out.value:
        raise OSError(f"CoCreateInstance failed: 0x{hr & 0xFFFFFFFF:08X}")
    return out.value


_oleaut32.SysFreeString.argtypes = [ctypes.c_void_p]
_oleaut32.SysFreeString.restype = None
_oleaut32.SysAllocString.argtypes = [ctypes.c_wchar_p]
_oleaut32.SysAllocString.restype = ctypes.c_void_p


def bstr_val(addr: int) -> Optional[str]:
    """Read BSTR at *addr*, free it, return Python str."""
    if not addr:
        return None
    try:
        return ctypes.wstring_at(addr)
    finally:
        _oleaut32.SysFreeString(addr)


def bstr_alloc(text: str) -> int:
    """Allocate a BSTR. Caller must free with ``bstr_free``."""
    return _oleaut32.SysAllocString(text)


def bstr_free(addr: int) -> None:
    _oleaut32.SysFreeString(addr)


_oleaut32.SafeArrayDestroy.argtypes = [ctypes.c_void_p]
_oleaut32.SafeArrayDestroy.restype = ctypes.c_long


class _SAFEARRAYBOUND(ctypes.Structure):
    _fields_ = [("cElements", ctypes.c_ulong), ("lLbound", ctypes.c_long)]


class _SAFEARRAY(ctypes.Structure):
    _fields_ = [
        ("cDims", ctypes.c_ushort),
        ("fFeatures", ctypes.c_ushort),
        ("cbElements", ctypes.c_ulong),
        ("cLocks", ctypes.c_ulong),
        ("pvData", ctypes.c_void_p),
        ("rgsabound", _SAFEARRAYBOUND * 1),
    ]


def sa_ints(psa_addr: int) -> Optional[list[int]]:
    """Convert SAFEARRAY(VT_I4)* to list[int] and destroy."""
    if not psa_addr:
        return None
    try:
        sa = ctypes.cast(psa_addr, ctypes.POINTER(_SAFEARRAY)).contents
        n = sa.rgsabound[0].cElements
        if n == 0:
            return []
        arr = (ctypes.c_int * n).from_address(sa.pvData)
        return list(arr)
    finally:
        _oleaut32.SafeArrayDestroy(psa_addr)
