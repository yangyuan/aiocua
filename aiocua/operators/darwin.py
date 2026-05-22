import asyncio
import ctypes
import ctypes.util
import base64
import hashlib
import re
from io import BytesIO
from typing import Optional
from PIL import Image
from aiocua.contracts.computer import (
    AxAction,
    AxNode,
    AxNodeBounds,
    AxNodeState,
    MonitorMetadata,
)
from aiocua.contracts.error import OperatorRuntimeException
from aiocua.operators.base import BaseCuaOperator

core = ctypes.cdll.LoadLibrary(ctypes.util.find_library("CoreGraphics"))
cf = ctypes.cdll.LoadLibrary(ctypes.util.find_library("CoreFoundation"))
app_services = ctypes.cdll.LoadLibrary(ctypes.util.find_library("ApplicationServices"))


class CGPoint(ctypes.Structure):
    _fields_ = [("x", ctypes.c_double), ("y", ctypes.c_double)]


class CGSize(ctypes.Structure):
    _fields_ = [("width", ctypes.c_double), ("height", ctypes.c_double)]


class CGRect(ctypes.Structure):
    _fields_ = [("origin", CGPoint), ("size", CGSize)]


class CFRange(ctypes.Structure):
    _fields_ = [("location", ctypes.c_long), ("length", ctypes.c_long)]


core.CGEventCreateMouseEvent.restype = ctypes.c_void_p
core.CGEventCreateMouseEvent.argtypes = [
    ctypes.c_void_p,
    ctypes.c_uint32,
    CGPoint,
    ctypes.c_uint32,
]
core.CGEventPost.restype = None
core.CGEventPost.argtypes = [ctypes.c_uint32, ctypes.c_void_p]
core.CFRelease.restype = None
core.CFRelease.argtypes = [ctypes.c_void_p]
core.CGEventCreate.restype = ctypes.c_void_p
core.CGEventCreate.argtypes = [ctypes.c_void_p]
core.CGEventGetLocation.restype = CGPoint
core.CGEventGetLocation.argtypes = [ctypes.c_void_p]
core.CGEventCreateKeyboardEvent.restype = ctypes.c_void_p
core.CGEventCreateKeyboardEvent.argtypes = [
    ctypes.c_void_p,
    ctypes.c_uint16,
    ctypes.c_bool,
]
core.CGEventCreateScrollWheelEvent.restype = ctypes.c_void_p
core.CGEventCreateScrollWheelEvent.argtypes = [
    ctypes.c_void_p,
    ctypes.c_int32,
    ctypes.c_uint32,
    ctypes.c_int32,
]
core.CGEventSetFlags.restype = None
core.CGEventSetFlags.argtypes = [ctypes.c_void_p, ctypes.c_uint64]

core.CGEventSetType.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
core.CGEventSetType.restype = None

core.CGEventSetIntegerValueField.argtypes = [
    ctypes.c_void_p,
    ctypes.c_int,
    ctypes.c_long,
]
core.CGEventSetIntegerValueField.restype = None

core.CGEventKeyboardSetUnicodeString.restype = None
core.CGEventKeyboardSetUnicodeString.argtypes = [
    ctypes.c_void_p,
    ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_uint16),
]

core.CGMainDisplayID.restype = ctypes.c_uint32
core.CGDisplayPixelsWide.restype = ctypes.c_size_t
core.CGDisplayPixelsWide.argtypes = [ctypes.c_uint32]
core.CGDisplayPixelsHigh.restype = ctypes.c_size_t
core.CGDisplayPixelsHigh.argtypes = [ctypes.c_uint32]
core.CGGetActiveDisplayList.restype = ctypes.c_int32
core.CGGetActiveDisplayList.argtypes = [
    ctypes.c_uint32,
    ctypes.POINTER(ctypes.c_uint32),
    ctypes.POINTER(ctypes.c_uint32),
]
core.CGDisplayCopyDisplayMode.restype = ctypes.c_void_p
core.CGDisplayCopyDisplayMode.argtypes = [ctypes.c_uint32]
core.CGDisplayModeGetPixelWidth.restype = ctypes.c_size_t
core.CGDisplayModeGetPixelWidth.argtypes = [ctypes.c_void_p]
core.CGDisplayModeGetPixelHeight.restype = ctypes.c_size_t
core.CGDisplayModeGetPixelHeight.argtypes = [ctypes.c_void_p]
core.CGWindowListCopyWindowInfo.restype = ctypes.c_void_p
core.CGWindowListCopyWindowInfo.argtypes = [ctypes.c_uint32, ctypes.c_uint32]

cf.CFRelease.restype = None
cf.CFRelease.argtypes = [ctypes.c_void_p]
cf.CFRetain.restype = ctypes.c_void_p
cf.CFRetain.argtypes = [ctypes.c_void_p]
cf.CFGetTypeID.restype = ctypes.c_ulong
cf.CFGetTypeID.argtypes = [ctypes.c_void_p]
cf.CFCopyDescription.restype = ctypes.c_void_p
cf.CFCopyDescription.argtypes = [ctypes.c_void_p]
cf.CFStringGetTypeID.restype = ctypes.c_ulong
cf.CFStringCreateWithCString.restype = ctypes.c_void_p
cf.CFStringCreateWithCString.argtypes = [
    ctypes.c_void_p,
    ctypes.c_char_p,
    ctypes.c_uint32,
]
cf.CFStringGetLength.restype = ctypes.c_long
cf.CFStringGetLength.argtypes = [ctypes.c_void_p]
cf.CFStringGetMaximumSizeForEncoding.restype = ctypes.c_long
cf.CFStringGetMaximumSizeForEncoding.argtypes = [ctypes.c_long, ctypes.c_uint32]
cf.CFStringGetCString.restype = ctypes.c_bool
cf.CFStringGetCString.argtypes = [
    ctypes.c_void_p,
    ctypes.c_char_p,
    ctypes.c_long,
    ctypes.c_uint32,
]
cf.CFArrayGetTypeID.restype = ctypes.c_ulong
cf.CFArrayGetCount.restype = ctypes.c_long
cf.CFArrayGetCount.argtypes = [ctypes.c_void_p]
cf.CFArrayGetValueAtIndex.restype = ctypes.c_void_p
cf.CFArrayGetValueAtIndex.argtypes = [ctypes.c_void_p, ctypes.c_long]
cf.CFBooleanGetTypeID.restype = ctypes.c_ulong
cf.CFBooleanGetValue.restype = ctypes.c_bool
cf.CFBooleanGetValue.argtypes = [ctypes.c_void_p]
cf.CFNumberGetTypeID.restype = ctypes.c_ulong
cf.CFNumberGetValue.restype = ctypes.c_bool
cf.CFNumberGetValue.argtypes = [
    ctypes.c_void_p,
    ctypes.c_int,
    ctypes.c_void_p,
]
cf.CFNumberIsFloatType.restype = ctypes.c_bool
cf.CFNumberIsFloatType.argtypes = [ctypes.c_void_p]
cf.CFDictionaryGetValue.restype = ctypes.c_void_p
cf.CFDictionaryGetValue.argtypes = [ctypes.c_void_p, ctypes.c_void_p]

app_services.AXIsProcessTrusted.restype = ctypes.c_bool
app_services.AXUIElementGetTypeID.restype = ctypes.c_ulong
app_services.AXUIElementCreateSystemWide.restype = ctypes.c_void_p
app_services.AXUIElementCreateApplication.restype = ctypes.c_void_p
app_services.AXUIElementCreateApplication.argtypes = [ctypes.c_int]
app_services.AXUIElementGetPid.restype = ctypes.c_int32
app_services.AXUIElementGetPid.argtypes = [
    ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_int),
]
app_services.AXUIElementCopyAttributeValue.restype = ctypes.c_int32
app_services.AXUIElementCopyAttributeValue.argtypes = [
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_void_p),
]
app_services.AXUIElementSetAttributeValue.restype = ctypes.c_int32
app_services.AXUIElementSetAttributeValue.argtypes = [
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.c_void_p,
]
app_services.AXUIElementIsAttributeSettable.restype = ctypes.c_int32
app_services.AXUIElementIsAttributeSettable.argtypes = [
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_bool),
]
app_services.AXUIElementCopyActionNames.restype = ctypes.c_int32
app_services.AXUIElementCopyActionNames.argtypes = [
    ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_void_p),
]
app_services.AXUIElementPerformAction.restype = ctypes.c_int32
app_services.AXUIElementPerformAction.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
app_services.AXValueGetTypeID.restype = ctypes.c_ulong
app_services.AXValueGetType.restype = ctypes.c_int32
app_services.AXValueGetType.argtypes = [ctypes.c_void_p]
app_services.AXValueGetValue.restype = ctypes.c_bool
app_services.AXValueGetValue.argtypes = [
    ctypes.c_void_p,
    ctypes.c_int32,
    ctypes.c_void_p,
]

kCGMouseEventClickState = 1

kCGEventLeftMouseDown = 1
kCGEventLeftMouseUp = 2
kCGEventRightMouseDown = 3
kCGEventRightMouseUp = 4
kCGEventMouseMoved = 5
kCGEventLeftMouseDragged = 6
kCGEventRightMouseDragged = 7
kCGEventScrollWheel = 22
kCGEventOtherMouseDown = 25
kCGEventOtherMouseUp = 26
kCGEventOtherMouseDragged = 27

kCGMouseButtonLeft = 0
kCGMouseButtonRight = 1
kCGMouseButtonCenter = 2

kCGScrollEventUnitPixel = 0
kCGScrollEventUnitLine = 1

kCGEventKeyDown = 10
kCGEventKeyUp = 11

kCGHIDEventTap = 0

# Keycode map (complete for standard US keyboard)
KEY_MAP = {
    "A": 0,
    "S": 1,
    "D": 2,
    "F": 3,
    "H": 4,
    "G": 5,
    "Z": 6,
    "X": 7,
    "C": 8,
    "V": 9,
    "B": 11,
    "Q": 12,
    "W": 13,
    "E": 14,
    "R": 15,
    "Y": 16,
    "T": 17,
    "1": 18,
    "2": 19,
    "3": 20,
    "4": 21,
    "6": 22,
    "5": 23,
    "EQUAL": 24,
    "9": 25,
    "7": 26,
    "MINUS": 27,
    "8": 28,
    "0": 29,
    "RIGHTBRACKET": 30,
    "O": 31,
    "U": 32,
    "LEFTBRACKET": 33,
    "I": 34,
    "P": 35,
    "RETURN": 36,
    "L": 37,
    "J": 38,
    "APOSTROPHE": 39,
    "K": 40,
    "SEMICOLON": 41,
    "BACKSLASH": 42,
    "COMMA": 43,
    "SLASH": 44,
    "N": 45,
    "M": 46,
    "PERIOD": 47,
    "TAB": 48,
    "SPACE": 49,
    "GRAVE": 50,
    "DELETE": 51,
    "ESC": 53,
    "CMD": 55,
    "SHIFT": 56,
    "CAPSLOCK": 57,
    "ALT": 58,
    "OPTION": 58,
    "CTRL": 59,
    "RIGHTSHIFT": 60,
    "RIGHTALT": 61,
    "RIGHTCTRL": 62,
    "FN": 63,
    "F17": 64,
    "KEYPADDECIMAL": 65,
    "KEYPADMULTIPLY": 67,
    "KEYPADPLUS": 69,
    "CLEAR": 71,
    "KEYPADDIVIDE": 75,
    "KEYPADENTER": 76,
    "KEYPADMINUS": 78,
    "F18": 79,
    "F19": 80,
    "KEYPADEQUAL": 81,
    "KEYPAD0": 82,
    "KEYPAD1": 83,
    "KEYPAD2": 84,
    "KEYPAD3": 85,
    "KEYPAD4": 86,
    "KEYPAD5": 87,
    "KEYPAD6": 88,
    "KEYPAD7": 89,
    "F20": 90,
    "KEYPAD8": 91,
    "KEYPAD9": 92,
    "F5": 96,
    "F6": 97,
    "F7": 98,
    "F3": 99,
    "F8": 100,
    "F9": 101,
    "F11": 103,
    "F13": 105,
    "F16": 106,
    "F14": 107,
    "F10": 109,
    "F12": 111,
    "F15": 113,
    "HELP": 114,
    "HOME": 115,
    "PAGEUP": 116,
    "FORWARDDELETE": 117,
    "F4": 118,
    "END": 119,
    "F2": 120,
    "PAGEDOWN": 121,
    "F1": 122,
    "LEFT": 123,
    "RIGHT": 124,
    "DOWN": 125,
    "UP": 126,
}

SLEEP_INTERVAL = 0.01
_MAX_NODES = 500

_AX_SUCCESS = 0
_AX_VALUE_CGPOINT = 1
_AX_VALUE_CGSIZE = 2
_AX_VALUE_CGRECT = 3
_AX_VALUE_CFRANGE = 4
_CF_STRING_ENCODING_UTF8 = 0x08000100
_CF_NUMBER_LONG_LONG = 11
_CF_NUMBER_DOUBLE = 13
_CG_WINDOW_LIST_ON_SCREEN_ONLY = 1
_CG_NULL_WINDOW_ID = 0

_CFSTRING_TYPE = cf.CFStringGetTypeID()
_CFARRAY_TYPE = cf.CFArrayGetTypeID()
_CFBOOLEAN_TYPE = cf.CFBooleanGetTypeID()
_CFNUMBER_TYPE = cf.CFNumberGetTypeID()
_AXUIELEMENT_TYPE = app_services.AXUIElementGetTypeID()
_AXVALUE_TYPE = app_services.AXValueGetTypeID()

_K_CF_BOOLEAN_TRUE = ctypes.c_void_p.in_dll(cf, "kCFBooleanTrue").value
_K_CF_BOOLEAN_FALSE = ctypes.c_void_p.in_dll(cf, "kCFBooleanFalse").value

_CFSTR_CACHE: dict[str, int] = {}

_AX_PRESS_ACTIONS = ("AXPress", "AXConfirm", "AXPick")
_AX_SCROLL_ACTIONS = {
    "up": "AXScrollUp",
    "down": "AXScrollDown",
    "left": "AXScrollLeft",
    "right": "AXScrollRight",
}
_AX_ID_MAX_PART_LENGTH = 48


def _get_mouse_pos() -> tuple[float, float]:
    event = core.CGEventCreate(None)
    pt = core.CGEventGetLocation(event)
    core.CFRelease(event)
    return pt.x, pt.y


def _cfstr(value: str) -> int:
    cached = _CFSTR_CACHE.get(value)
    if cached is not None:
        return cached
    created = cf.CFStringCreateWithCString(
        None, value.encode("utf-8"), _CF_STRING_ENCODING_UTF8
    )
    if not created:
        raise OperatorRuntimeException(f"Cannot create CFString for '{value}'")
    _CFSTR_CACHE[value] = created
    return created


def _cf_release(value: Optional[int]) -> None:
    if value:
        cf.CFRelease(value)


def _cf_retain(value: int) -> int:
    return cf.CFRetain(value)


def _cf_type(value: Optional[int]) -> Optional[int]:
    if not value:
        return None
    return cf.CFGetTypeID(value)


def _cf_to_str(value: Optional[int]) -> Optional[str]:
    if not value or _cf_type(value) != _CFSTRING_TYPE:
        return None
    length = cf.CFStringGetLength(value)
    max_size = cf.CFStringGetMaximumSizeForEncoding(length, _CF_STRING_ENCODING_UTF8) + 1
    buf = ctypes.create_string_buffer(max_size)
    if not cf.CFStringGetCString(value, buf, max_size, _CF_STRING_ENCODING_UTF8):
        return None
    return buf.value.decode("utf-8", errors="replace")


def _cf_description(value: Optional[int]) -> Optional[str]:
    if not value:
        return None
    description = cf.CFCopyDescription(value)
    try:
        return _cf_to_str(description)
    finally:
        _cf_release(description)


def _cf_to_int(value: Optional[int]) -> Optional[int]:
    if not value or _cf_type(value) != _CFNUMBER_TYPE:
        return None
    out = ctypes.c_longlong()
    if not cf.CFNumberGetValue(value, _CF_NUMBER_LONG_LONG, ctypes.byref(out)):
        return None
    return int(out.value)


def _format_number(value: float) -> str:
    text = f"{value:.6f}".rstrip("0").rstrip(".")
    return "0" if text == "-0" else text


def _cf_number_to_str(value: Optional[int]) -> Optional[str]:
    if not value or _cf_type(value) != _CFNUMBER_TYPE:
        return None
    if cf.CFNumberIsFloatType(value):
        out = ctypes.c_double()
        if cf.CFNumberGetValue(value, _CF_NUMBER_DOUBLE, ctypes.byref(out)):
            return _format_number(out.value)
    int_value = _cf_to_int(value)
    return None if int_value is None else str(int_value)


def _cf_to_bool(value: Optional[int]) -> Optional[bool]:
    if not value:
        return None
    value_type = _cf_type(value)
    if value_type == _CFBOOLEAN_TYPE:
        return bool(cf.CFBooleanGetValue(value))
    if value_type == _CFNUMBER_TYPE:
        int_value = _cf_to_int(value)
        return None if int_value is None else bool(int_value)
    return None


def _ax_copy_attr(el: int, name: str) -> Optional[int]:
    out = ctypes.c_void_p()
    err = app_services.AXUIElementCopyAttributeValue(
        el, _cfstr(name), ctypes.byref(out)
    )
    return out.value if err == _AX_SUCCESS and out.value else None


def _ax_pid(el: int) -> Optional[int]:
    pid = ctypes.c_int()
    err = app_services.AXUIElementGetPid(el, ctypes.byref(pid))
    return pid.value if err == _AX_SUCCESS else None


def _ax_string_attr(el: int, *names: str) -> Optional[str]:
    for name in names:
        value = _ax_copy_attr(el, name)
        try:
            result = _cf_to_str(value)
            if result:
                return result
        finally:
            _cf_release(value)
    return None


def _ax_bool_attr(el: int, name: str, default: Optional[bool] = None) -> Optional[bool]:
    value = _ax_copy_attr(el, name)
    try:
        result = _cf_to_bool(value)
        return default if result is None else result
    finally:
        _cf_release(value)


def _ax_is_attr_settable(el: int, name: str) -> bool:
    settable = ctypes.c_bool(False)
    err = app_services.AXUIElementIsAttributeSettable(
        el, _cfstr(name), ctypes.byref(settable)
    )
    return err == _AX_SUCCESS and bool(settable.value)


def _ax_set_bool_attr(el: int, name: str, value: bool) -> bool:
    cf_bool = _K_CF_BOOLEAN_TRUE if value else _K_CF_BOOLEAN_FALSE
    err = app_services.AXUIElementSetAttributeValue(el, _cfstr(name), cf_bool)
    return err == _AX_SUCCESS


def _ax_set_string_attr(el: int, name: str, value: str) -> bool:
    cf_value = cf.CFStringCreateWithCString(
        None, value.encode("utf-8"), _CF_STRING_ENCODING_UTF8
    )
    if not cf_value:
        return False
    try:
        err = app_services.AXUIElementSetAttributeValue(el, _cfstr(name), cf_value)
        return err == _AX_SUCCESS
    finally:
        _cf_release(cf_value)


def _ax_action_names(el: int) -> set[str]:
    actions = ctypes.c_void_p()
    err = app_services.AXUIElementCopyActionNames(el, ctypes.byref(actions))
    if err != _AX_SUCCESS or not actions.value:
        return set()

    result: set[str] = set()
    try:
        if _cf_type(actions.value) != _CFARRAY_TYPE:
            return result
        count = cf.CFArrayGetCount(actions.value)
        for index in range(count):
            item = cf.CFArrayGetValueAtIndex(actions.value, index)
            text = _cf_to_str(item)
            if text:
                result.add(text)
    finally:
        _cf_release(actions.value)
    return result


def _ax_perform_action(el: int, action: str) -> bool:
    err = app_services.AXUIElementPerformAction(el, _cfstr(action))
    return err == _AX_SUCCESS


def _ax_perform_first_action(el: int, actions: tuple[str, ...]) -> bool:
    supported = _ax_action_names(el)
    for action in actions:
        if action in supported and _ax_perform_action(el, action):
            return True
    return False


def _ax_array_elements(value: Optional[int]) -> list[int]:
    if not value or _cf_type(value) != _CFARRAY_TYPE:
        return []
    result: list[int] = []
    count = cf.CFArrayGetCount(value)
    for index in range(count):
        item = cf.CFArrayGetValueAtIndex(value, index)
        if item and _cf_type(item) == _AXUIELEMENT_TYPE:
            result.append(_cf_retain(item))
    return result


def _ax_elements_attr(el: int, name: str) -> list[int]:
    value = _ax_copy_attr(el, name)
    try:
        return _ax_array_elements(value)
    finally:
        _cf_release(value)


def _ax_bounds(el: int) -> Optional[AxNodeBounds]:
    position_value = _ax_copy_attr(el, "AXPosition")
    size_value = _ax_copy_attr(el, "AXSize")
    try:
        if (
            not position_value
            or not size_value
            or _cf_type(position_value) != _AXVALUE_TYPE
            or _cf_type(size_value) != _AXVALUE_TYPE
        ):
            return None

        point = CGPoint()
        size = CGSize()
        if not app_services.AXValueGetValue(
            position_value, _AX_VALUE_CGPOINT, ctypes.byref(point)
        ):
            return None
        if not app_services.AXValueGetValue(
            size_value, _AX_VALUE_CGSIZE, ctypes.byref(size)
        ):
            return None
        return AxNodeBounds(
            x=round(point.x),
            y=round(point.y),
            w=max(0, round(size.width)),
            h=max(0, round(size.height)),
        )
    finally:
        _cf_release(position_value)
        _cf_release(size_value)


def _ax_center(el: int) -> tuple[int, int]:
    bounds = _ax_bounds(el)
    if bounds is None:
        raise OperatorRuntimeException("Cannot get accessibility element bounds")
    return bounds.x + bounds.w // 2, bounds.y + bounds.h // 2


def _ax_id_part(value: Optional[object]) -> str:
    if value is None or value == "":
        return "none"
    text = str(value).strip().lower()
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"[^a-z0-9_.:-]+", "", text)
    return (text or "none")[:_AX_ID_MAX_PART_LENGTH]


def _ax_id_bounds(bounds: Optional[AxNodeBounds]) -> str:
    if bounds is None:
        return "bounds:none"
    return f"bounds:{bounds.x},{bounds.y},{bounds.w},{bounds.h}"


def _ax_stable_id(
    el: int,
    parent_id: Optional[str],
    role: str,
    subrole: Optional[str],
    name: str,
    description: str,
    bounds: Optional[AxNodeBounds],
) -> str:
    pid = _ax_pid(el)
    identifier = _ax_string_attr(el, "AXIdentifier")
    role_part = _ax_id_part(role)
    name_part = _ax_id_part(identifier or name or description or subrole)
    basis = "|".join(
        (
            "macos",
            f"pid:{pid}" if pid is not None else "pid:none",
            f"parent:{parent_id or 'root'}",
            f"role:{role}",
            f"subrole:{subrole or ''}",
            f"identifier:{identifier or ''}",
            f"name:{name}",
            f"description:{description}",
            _ax_id_bounds(bounds),
        )
    )
    digest = hashlib.sha1(basis.encode("utf-8")).hexdigest()[:10]
    return f"macos:{role_part}:{name_part}:{digest}"


def _ax_node_name(el: int) -> str:
    return _ax_string_attr(el, "AXTitle") or ""


def _ax_node_description(el: int) -> str:
    return _ax_string_attr(el, "AXDescription", "AXHelp") or ""


def _cf_dict_value(dictionary: int, key: str) -> Optional[int]:
    value = cf.CFDictionaryGetValue(dictionary, _cfstr(key))
    return value if value else None


def _cf_dict_int(dictionary: int, key: str) -> Optional[int]:
    return _cf_to_int(_cf_dict_value(dictionary, key))


def _cf_dict_bool(dictionary: int, key: str) -> Optional[bool]:
    return _cf_to_bool(_cf_dict_value(dictionary, key))


def _visible_window_owner_pids() -> list[int]:
    info = core.CGWindowListCopyWindowInfo(
        _CG_WINDOW_LIST_ON_SCREEN_ONLY, _CG_NULL_WINDOW_ID
    )
    if not info:
        return []

    seen: set[int] = set()
    pids: list[int] = []
    try:
        if _cf_type(info) != _CFARRAY_TYPE:
            return pids
        count = cf.CFArrayGetCount(info)
        for index in range(count):
            window = cf.CFArrayGetValueAtIndex(info, index)
            if not window:
                continue
            layer = _cf_dict_int(window, "kCGWindowLayer")
            if layer not in (None, 0):
                continue
            onscreen = _cf_dict_bool(window, "kCGWindowIsOnscreen")
            if onscreen is False:
                continue
            pid = _cf_dict_int(window, "kCGWindowOwnerPID")
            if pid is None or pid in seen:
                continue
            seen.add(pid)
            pids.append(pid)
    finally:
        _cf_release(info)
    return pids


def _application_windows(pid: int) -> list[int]:
    app = app_services.AXUIElementCreateApplication(pid)
    if not app:
        return []
    try:
        return _ax_elements_attr(app, "AXWindows")
    finally:
        _cf_release(app)


def _point_to_str(point: CGPoint) -> str:
    return f"x={_format_number(point.x)},y={_format_number(point.y)}"


def _size_to_str(size: CGSize) -> str:
    return f"width={_format_number(size.width)},height={_format_number(size.height)}"


def _ax_value_to_str(value: int) -> Optional[str]:
    value_type = app_services.AXValueGetType(value)
    if value_type == _AX_VALUE_CGPOINT:
        point = CGPoint()
        if app_services.AXValueGetValue(value, value_type, ctypes.byref(point)):
            return _point_to_str(point)
    if value_type == _AX_VALUE_CGSIZE:
        size = CGSize()
        if app_services.AXValueGetValue(value, value_type, ctypes.byref(size)):
            return _size_to_str(size)
    if value_type == _AX_VALUE_CGRECT:
        rect = CGRect()
        if app_services.AXValueGetValue(value, value_type, ctypes.byref(rect)):
            return f"{_point_to_str(rect.origin)},{_size_to_str(rect.size)}"
    if value_type == _AX_VALUE_CFRANGE:
        range_value = CFRange()
        if app_services.AXValueGetValue(value, value_type, ctypes.byref(range_value)):
            return f"location={range_value.location},length={range_value.length}"
    return _cf_description(value)


def _ax_value_attr(el: int, name: str) -> Optional[str]:
    value = _ax_copy_attr(el, name)
    try:
        if not value:
            return None
        value_type = _cf_type(value)
        if value_type == _CFSTRING_TYPE:
            return _cf_to_str(value)
        if value_type == _CFBOOLEAN_TYPE:
            bool_value = _cf_to_bool(value)
            return None if bool_value is None else str(bool_value).lower()
        if value_type == _CFNUMBER_TYPE:
            return _cf_number_to_str(value)
        if value_type == _AXVALUE_TYPE:
            return _ax_value_to_str(value)
        return _cf_description(value)
    finally:
        _cf_release(value)


def _ax_allowed_actions(
    el: int, supported_actions: set[str], bounds: Optional[AxNodeBounds]
) -> list[AxAction]:
    actions: list[AxAction] = []
    enabled = _ax_bool_attr(el, "AXEnabled", default=True)
    can_set_expanded = _ax_is_attr_settable(el, "AXExpanded")

    if enabled and (
        any(action in supported_actions for action in _AX_PRESS_ACTIONS)
        or _ax_is_attr_settable(el, "AXSelected")
        or bounds is not None
    ):
        actions.append(AxAction.CLICK)
    if _ax_is_attr_settable(el, "AXFocused") or "AXRaise" in supported_actions:
        actions.append(AxAction.FOCUS)
    if _ax_is_attr_settable(el, "AXValue"):
        actions.append(AxAction.TYPE)
    if any(action in supported_actions for action in _AX_SCROLL_ACTIONS.values()) or (
        "AXScrollToVisible" in supported_actions
    ):
        actions.append(AxAction.SCROLL)
    if can_set_expanded or "AXShowMenu" in supported_actions:
        actions.append(AxAction.EXPAND)
    if can_set_expanded:
        actions.append(AxAction.COLLAPSE)
    if _ax_is_attr_settable(el, "AXSelected") or "AXPress" in supported_actions:
        actions.append(AxAction.SELECT)
    return actions


class DarwinCuaOperator(BaseCuaOperator):
    async def move(self, x: int, y: int) -> None:
        event = core.CGEventCreateMouseEvent(
            None, kCGEventMouseMoved, CGPoint(x, y), kCGMouseButtonLeft
        )
        core.CGEventPost(kCGHIDEventTap, event)
        core.CFRelease(event)

    async def click(
        self, button: str = "left", x: Optional[int] = None, y: Optional[int] = None
    ) -> None:
        if x is not None and y is not None:
            await self.move(x, y)
            await asyncio.sleep(SLEEP_INTERVAL)
        btn_map = {
            "left": kCGMouseButtonLeft,
            "right": kCGMouseButtonRight,
            "wheel": kCGMouseButtonCenter,
            "middle": kCGMouseButtonCenter,
            "back": 3,
            "forward": 4,
        }
        if button not in btn_map:
            raise OperatorRuntimeException(f"Unknown button: {button}")
        btn = btn_map[button]
        if x is not None and y is not None:
            pt = CGPoint(x, y)
        else:
            loc = _get_mouse_pos()
            pt = CGPoint(*loc)
        if btn == kCGMouseButtonLeft:
            events = [kCGEventLeftMouseDown, kCGEventLeftMouseUp]
        elif btn == kCGMouseButtonRight:
            events = [kCGEventRightMouseDown, kCGEventRightMouseUp]
        elif btn == kCGMouseButtonCenter:
            events = [kCGEventOtherMouseDown, kCGEventOtherMouseUp]
        else:
            events = [kCGEventOtherMouseDown, kCGEventOtherMouseUp]
        for t in events:
            await asyncio.sleep(SLEEP_INTERVAL)
            event = core.CGEventCreateMouseEvent(None, t, pt, btn)
            core.CGEventPost(kCGHIDEventTap, event)
            core.CFRelease(event)

    async def double_click(
        self, x: Optional[int] = None, y: Optional[int] = None
    ) -> None:
        if x is not None and y is not None:
            await self.move(x, y)
            await asyncio.sleep(SLEEP_INTERVAL)
            pt = CGPoint(x, y)
        else:
            loc = _get_mouse_pos()
            pt = CGPoint(*loc)
        event_dclick = core.CGEventCreateMouseEvent(
            None, kCGEventLeftMouseDown, pt, kCGMouseButtonLeft
        )
        core.CGEventSetIntegerValueField(event_dclick, kCGMouseEventClickState, 1)
        core.CGEventPost(kCGHIDEventTap, event_dclick)
        core.CGEventSetType(event_dclick, kCGEventLeftMouseUp)
        core.CGEventPost(kCGHIDEventTap, event_dclick)
        await asyncio.sleep(SLEEP_INTERVAL)
        core.CGEventSetIntegerValueField(event_dclick, kCGMouseEventClickState, 2)
        core.CGEventSetType(event_dclick, kCGEventLeftMouseDown)
        core.CGEventPost(kCGHIDEventTap, event_dclick)
        core.CGEventSetType(event_dclick, kCGEventLeftMouseUp)
        core.CGEventPost(kCGHIDEventTap, event_dclick)
        await asyncio.sleep(SLEEP_INTERVAL)
        core.CFRelease(event_dclick)

    async def drag(self, path: list[tuple[int, int]]) -> None:
        if not path or len(path) < 2:
            raise OperatorRuntimeException("Path must contain at least two points")
        pt_start = CGPoint(*path[0])
        event = core.CGEventCreateMouseEvent(
            None, kCGEventLeftMouseDown, pt_start, kCGMouseButtonLeft
        )
        core.CGEventPost(kCGHIDEventTap, event)
        core.CFRelease(event)
        await asyncio.sleep(SLEEP_INTERVAL)
        for point in path[1:-1]:
            pt = CGPoint(*point)
            event = core.CGEventCreateMouseEvent(
                None, kCGEventLeftMouseDragged, pt, kCGMouseButtonLeft
            )
            core.CGEventPost(kCGHIDEventTap, event)
            core.CFRelease(event)
            await asyncio.sleep(SLEEP_INTERVAL)
        pt_end = CGPoint(*path[-1])
        event = core.CGEventCreateMouseEvent(
            None, kCGEventLeftMouseDragged, pt_end, kCGMouseButtonLeft
        )
        core.CGEventPost(kCGHIDEventTap, event)
        core.CFRelease(event)
        await asyncio.sleep(SLEEP_INTERVAL)
        event = core.CGEventCreateMouseEvent(
            None, kCGEventLeftMouseUp, pt_end, kCGMouseButtonLeft
        )
        core.CGEventPost(kCGHIDEventTap, event)
        core.CFRelease(event)

    async def key_press(self, keys: list[str]) -> None:
        MODIFIER_FLAGS = {
            "SHIFT": 0x20000,
            "CTRL": 0x40000,
            "ALT": 0x80000,
            "CMD": 0x100000,
            "OPTION": 0x80000,
        }

        keys_upper = [k.upper() for k in keys]
        modifiers = [k for k in keys_upper if k in MODIFIER_FLAGS]
        normal_keys = [k for k in keys_upper if k not in MODIFIER_FLAGS]

        if not normal_keys:
            keycodes = [KEY_MAP[k] for k in modifiers if k in KEY_MAP]
            for keycode in keycodes:
                event_down = core.CGEventCreateKeyboardEvent(None, keycode, True)
                core.CGEventPost(kCGHIDEventTap, event_down)
                core.CFRelease(event_down)
            await asyncio.sleep(SLEEP_INTERVAL)
            for keycode in reversed(keycodes):
                event_up = core.CGEventCreateKeyboardEvent(None, keycode, False)
                core.CGEventPost(kCGHIDEventTap, event_up)
                core.CFRelease(event_up)
            return

        flags = 0
        for mod in modifiers:
            flags |= MODIFIER_FLAGS[mod]

        modifier_keycodes = []
        for mod in modifiers:
            if mod in KEY_MAP:
                keycode = KEY_MAP[mod]
                modifier_keycodes.append(keycode)
                event_down = core.CGEventCreateKeyboardEvent(None, keycode, True)
                core.CGEventPost(kCGHIDEventTap, event_down)
                core.CFRelease(event_down)

        for key in normal_keys:
            if key in KEY_MAP:
                keycode = KEY_MAP[key]
            else:
                raise OperatorRuntimeException(f"Unknown key: {key}")

            event_down = core.CGEventCreateKeyboardEvent(None, keycode, True)
            if flags:
                core.CGEventSetFlags(event_down, flags)
            core.CGEventPost(kCGHIDEventTap, event_down)
            core.CFRelease(event_down)

            event_up = core.CGEventCreateKeyboardEvent(None, keycode, False)
            if flags:
                core.CGEventSetFlags(event_up, flags)
            core.CGEventPost(kCGHIDEventTap, event_up)
            core.CFRelease(event_up)

        await asyncio.sleep(SLEEP_INTERVAL)

        for keycode in reversed(modifier_keycodes):
            event_up = core.CGEventCreateKeyboardEvent(None, keycode, False)
            core.CGEventPost(kCGHIDEventTap, event_up)
            core.CFRelease(event_up)

    async def type_text(self, text: str) -> None:
        for char in text:
            event_down = core.CGEventCreateKeyboardEvent(None, 0, True)
            utf16 = char.encode("utf-16-le")
            length = len(utf16) // 2
            buf = (ctypes.c_uint16 * length).from_buffer_copy(utf16)
            core.CGEventKeyboardSetUnicodeString(event_down, length, buf)
            core.CGEventPost(kCGHIDEventTap, event_down)
            core.CFRelease(event_down)
            await asyncio.sleep(SLEEP_INTERVAL)
            event_up = core.CGEventCreateKeyboardEvent(None, 0, False)
            core.CGEventKeyboardSetUnicodeString(event_up, length, buf)
            core.CGEventPost(kCGHIDEventTap, event_up)
            core.CFRelease(event_up)
            await asyncio.sleep(SLEEP_INTERVAL)

    async def scroll(
        self,
        scroll_x: int,
        scroll_y: int,
        x: Optional[int] = None,
        y: Optional[int] = None,
    ) -> None:
        if x is not None and y is not None:
            await self.move(x, y)
            await asyncio.sleep(SLEEP_INTERVAL)
        event = None
        if scroll_y != 0 and scroll_x != 0:
            event = core.CGEventCreateScrollWheelEvent(
                None, kCGScrollEventUnitLine, 2, int(scroll_y), int(scroll_x)
            )
        elif scroll_y != 0:
            event = core.CGEventCreateScrollWheelEvent(
                None, kCGScrollEventUnitLine, 1, int(scroll_y)
            )
        elif scroll_x != 0:
            event = core.CGEventCreateScrollWheelEvent(
                None, kCGScrollEventUnitLine, 2, 0, int(scroll_x)
            )
        if event:
            core.CGEventPost(kCGHIDEventTap, event)
            core.CFRelease(event)

    async def screenshot(self) -> str:
        from PIL import ImageGrab

        img = ImageGrab.grab()
        width, height = await self.dimensions()
        if img.size != (width, height):
            img = img.resize((width, height), Image.LANCZOS)
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
        return img_str

    async def dimensions(self) -> tuple[int, int]:
        display_id = core.CGMainDisplayID()
        width = core.CGDisplayPixelsWide(display_id)
        height = core.CGDisplayPixelsHigh(display_id)
        return width, height

    async def monitors(self) -> list[MonitorMetadata]:
        max_displays = 16
        display_ids = (ctypes.c_uint32 * max_displays)()
        display_count = ctypes.c_uint32()

        core.CGGetActiveDisplayList(
            max_displays, display_ids, ctypes.byref(display_count)
        )

        main_id = core.CGMainDisplayID()

        result: list[MonitorMetadata] = []
        for i in range(display_count.value):
            did = display_ids[i]

            logical_w = core.CGDisplayPixelsWide(did)
            logical_h = core.CGDisplayPixelsHigh(did)

            scale = 1.0
            mode = core.CGDisplayCopyDisplayMode(did)
            if mode:
                pixel_w = core.CGDisplayModeGetPixelWidth(mode)
                scale = pixel_w / logical_w if logical_w else 1.0
                core.CFRelease(mode)

            result.append(
                MonitorMetadata(
                    id=int(did),
                    width=logical_w,
                    height=logical_h,
                    scale=scale,
                    primary=did == main_id,
                )
            )
        return result

    async def wait(self) -> None:
        await asyncio.sleep(1)

    def _ax_init(self) -> None:
        if hasattr(self, "_ax_cache"):
            return
        if not app_services.AXIsProcessTrusted():
            raise OperatorRuntimeException(
                "macOS Accessibility permission is required for AX operations. "
                "Enable it in System Settings > Privacy & Security > Accessibility."
            )
        self._ax_cache: dict[str, int] = {}
        self._ax_seen_ids: set[str] = set()
        self._ax_node_count = 0

    def _ax_element(self, node_id: str) -> int:
        el = self._ax_cache.get(node_id)
        if el is None:
            raise OperatorRuntimeException(
                f"Node '{node_id}' not found. Call axtree() first."
            )
        return el

    def _ax_clear_cache(self) -> None:
        for el in self._ax_cache.values():
            _cf_release(el)
        self._ax_cache.clear()

    def _ax_desktop_children(self, systemwide: int) -> list[int]:
        children: list[int] = []
        for pid in _visible_window_owner_pids():
            children.extend(_application_windows(pid))
        if children:
            return children

        focused_app = _ax_copy_attr(systemwide, "AXFocusedApplication")
        try:
            if focused_app and _cf_type(focused_app) == _AXUIELEMENT_TYPE:
                return _ax_elements_attr(focused_app, "AXWindows")
        finally:
            _cf_release(focused_app)

        focused_window = _ax_copy_attr(systemwide, "AXFocusedWindow")
        if focused_window and _cf_type(focused_window) == _AXUIELEMENT_TYPE:
            return [focused_window]
        _cf_release(focused_window)
        return []

    def _ax_children(self, el: int, raw_role: Optional[str]) -> list[int]:
        if raw_role == "AXSystemWide":
            children = _ax_elements_attr(el, "AXChildren")
            return children or self._ax_desktop_children(el)
        if raw_role == "AXApplication":
            children = _ax_elements_attr(el, "AXWindows")
            if children:
                return children
        return _ax_elements_attr(el, "AXChildren")

    def _ax_states(self, el: int, role: str) -> list[AxNodeState]:
        states: list[AxNodeState] = [AxNodeState.VISIBLE]
        enabled = _ax_bool_attr(el, "AXEnabled", default=True)
        if enabled:
            states.append(AxNodeState.ENABLED)
        if _ax_bool_attr(el, "AXFocused", default=False):
            states.append(AxNodeState.FOCUSED)
        if enabled and (
            _ax_is_attr_settable(el, "AXFocused")
            or bool(_ax_action_names(el))
        ):
            states.append(AxNodeState.FOCUSABLE)
        if _ax_bool_attr(el, "AXSelected", default=False):
            states.append(AxNodeState.SELECTED)
        if _ax_bool_attr(el, "AXExpanded", default=False):
            states.append(AxNodeState.EXPANDED)
        return states

    def _ax_node_id(
        self, el: int, parent_id: Optional[str], forced_id: Optional[str] = None
    ) -> tuple[str, str, str, str, Optional[AxNodeBounds]]:
        raw_role = _ax_string_attr(el, "AXRole") or "AXUnknown"
        subrole = _ax_string_attr(el, "AXSubrole")
        name = _ax_node_name(el)
        description = _ax_node_description(el)
        bounds = _ax_bounds(el)
        node_id = forced_id or _ax_stable_id(
            el, parent_id, raw_role, subrole, name, description, bounds
        )
        if node_id in self._ax_seen_ids:
            suffix = 2
            base_id = node_id
            while node_id in self._ax_seen_ids:
                node_id = f"{base_id}:{suffix}"
                suffix += 1
        self._ax_seen_ids.add(node_id)
        return node_id, raw_role, name, description, bounds

    def _ax_walk(
        self,
        el: int,
        parent_id: Optional[str],
        depth: int,
        max_depth: int,
        forced_id: Optional[str] = None,
    ) -> Optional[AxNode]:
        if depth > max_depth or self._ax_node_count >= _MAX_NODES or not el:
            return None
        if _ax_bool_attr(el, "AXHidden", default=False):
            return None

        node_id, raw_role, name, description, bounds = self._ax_node_id(
            el, parent_id, forced_id
        )
        subrole = _ax_string_attr(el, "AXSubrole")
        supported_actions = _ax_action_names(el)
        self._ax_node_count += 1
        self._ax_cache[node_id] = _cf_retain(el)

        children: list[AxNode] = []
        child_refs = self._ax_children(el, raw_role)
        for child in child_refs:
            try:
                child_node = self._ax_walk(
                    child, node_id, depth + 1, max_depth
                )
                if child_node is not None:
                    children.append(child_node)
            finally:
                _cf_release(child)

        return AxNode(
            id=node_id,
            role=raw_role,
            secondary_role=subrole,
            name=name,
            bounds=bounds,
            states=self._ax_states(el, raw_role),
            children=children,
            description=description,
            value=_ax_value_attr(el, "AXValue"),
            allowed_actions=_ax_allowed_actions(el, supported_actions, bounds),
        )

    async def axtree(
        self, root_node_id: Optional[str] = None, max_depth: int = 8
    ) -> AxNode:
        self._ax_init()
        if max_depth < 0:
            raise OperatorRuntimeException("max_depth must be >= 0")

        if root_node_id:
            root = _cf_retain(self._ax_element(root_node_id))
            forced_id = root_node_id
            parent_id = None
        else:
            root = app_services.AXUIElementCreateSystemWide()
            forced_id = None
            parent_id = None
            if not root:
                raise OperatorRuntimeException("Failed to create macOS AX root element")

        self._ax_clear_cache()
        self._ax_node_count = 0
        self._ax_seen_ids = set()
        try:
            tree = self._ax_walk(root, parent_id, 0, max_depth, forced_id)
        finally:
            _cf_release(root)

        if tree is None:
            return AxNode(id="", role="empty")
        return tree

    async def ax_click(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)

        if _ax_perform_first_action(el, _AX_PRESS_ACTIONS):
            return
        if _ax_is_attr_settable(el, "AXSelected") and _ax_set_bool_attr(
            el, "AXSelected", True
        ):
            return

        x, y = _ax_center(el)
        await self.click(button="left", x=x, y=y)

    async def ax_double_click(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        x, y = _ax_center(el)
        await self.double_click(x=x, y=y)

    async def ax_type(self, text: str, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        if _ax_is_attr_settable(el, "AXValue") and _ax_set_string_attr(
            el, "AXValue", text
        ):
            return
        await self.ax_focus(node_id)
        await asyncio.sleep(0.05)
        await self.type_text(text)

    async def ax_scroll(self, scroll_x: int, scroll_y: int, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)

        actions: list[str] = []
        if scroll_y > 0:
            actions.append(_AX_SCROLL_ACTIONS["down"])
        elif scroll_y < 0:
            actions.append(_AX_SCROLL_ACTIONS["up"])
        if scroll_x > 0:
            actions.append(_AX_SCROLL_ACTIONS["right"])
        elif scroll_x < 0:
            actions.append(_AX_SCROLL_ACTIONS["left"])

        supported = _ax_action_names(el)
        if actions and all(action in supported for action in actions):
            if all(_ax_perform_action(el, action) for action in actions):
                return

        _ax_perform_action(el, "AXScrollToVisible")
        x, y = _ax_center(el)
        await self.scroll(scroll_x, scroll_y, x=x, y=y)

    async def ax_focus(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        _ax_perform_action(el, "AXRaise")
        if not _ax_set_bool_attr(el, "AXFocused", True):
            x, y = _ax_center(el)
            await self.click(button="left", x=x, y=y)

    async def ax_expand(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        if _ax_is_attr_settable(el, "AXExpanded") and _ax_set_bool_attr(
            el, "AXExpanded", True
        ):
            return
        if _ax_perform_action(el, "AXShowMenu") or _ax_perform_action(el, "AXPress"):
            return
        raise OperatorRuntimeException(
            f"Node '{node_id}' does not support expand/collapse."
        )

    async def ax_collapse(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        if _ax_is_attr_settable(el, "AXExpanded") and _ax_set_bool_attr(
            el, "AXExpanded", False
        ):
            return
        raise OperatorRuntimeException(
            f"Node '{node_id}' does not support expand/collapse."
        )

    async def ax_select(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        if _ax_is_attr_settable(el, "AXSelected") and _ax_set_bool_attr(
            el, "AXSelected", True
        ):
            return
        if _ax_perform_action(el, "AXPress"):
            return
        raise OperatorRuntimeException(f"Node '{node_id}' does not support selection.")
