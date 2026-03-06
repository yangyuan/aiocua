import asyncio
import ctypes
import base64
from io import BytesIO
from ctypes import wintypes
from typing import Optional, Union
from PIL import ImageGrab, Image
from aiocua.contracts.computer import AxNode, AxNodeBounds, AxNodeState, MonitorMetadata
from aiocua.contracts.error import OperatorRuntimeException
from aiocua.helpers import com
from aiocua.operators.base import BaseCuaOperator


user32 = ctypes.windll.user32


try:
    # PROCESS_PER_MONITOR_DPI_AWARE = 2
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except (OSError, AttributeError):
    # Fallback for Windows 7 / older
    try:
        user32.SetProcessDPIAware()
    except AttributeError:
        pass

SLEEP_INTERVAL = 0.01

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_WHEEL = 0x0800
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_HWHEEL = 0x01000

KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
VK_SHIFT = 0x10

VK_MAP = {
    "LBUTTON": 0x01,
    "RBUTTON": 0x02,
    "CANCEL": 0x03,
    "MBUTTON": 0x04,
    "XBUTTON1": 0x05,
    "XBUTTON2": 0x06,
    "BACK": 0x08,
    "TAB": 0x09,
    "CLEAR": 0x0C,
    "RETURN": 0x0D,
    "SHIFT": 0x10,
    "CONTROL": 0x11,
    "MENU": 0x12,
    "PAUSE": 0x13,
    "CAPITAL": 0x14,
    "KANA": 0x15,
    "HANGUL": 0x15,
    "IME_ON": 0x16,
    "JUNJA": 0x17,
    "FINAL": 0x18,
    "HANJA": 0x19,
    "KANJI": 0x19,
    "IME_OFF": 0x1A,
    "ESCAPE": 0x1B,
    "CONVERT": 0x1C,
    "NONCONVERT": 0x1D,
    "ACCEPT": 0x1E,
    "MODECHANGE": 0x1F,
    "SPACE": 0x20,
    "PRIOR": 0x21,
    "NEXT": 0x22,
    "END": 0x23,
    "HOME": 0x24,
    "LEFT": 0x25,
    "UP": 0x26,
    "RIGHT": 0x27,
    "DOWN": 0x28,
    "SELECT": 0x29,
    "PRINT": 0x2A,
    "EXECUTE": 0x2B,
    "SNAPSHOT": 0x2C,
    "INSERT": 0x2D,
    "DELETE": 0x2E,
    "HELP": 0x2F,
    "LWIN": 0x5B,
    "RWIN": 0x5C,
    "APPS": 0x5D,
    "SLEEP": 0x5F,
    "NUMPAD0": 0x60,
    "NUMPAD1": 0x61,
    "NUMPAD2": 0x62,
    "NUMPAD3": 0x63,
    "NUMPAD4": 0x64,
    "NUMPAD5": 0x65,
    "NUMPAD6": 0x66,
    "NUMPAD7": 0x67,
    "NUMPAD8": 0x68,
    "NUMPAD9": 0x69,
    "MULTIPLY": 0x6A,
    "ADD": 0x6B,
    "SEPARATOR": 0x6C,
    "SUBTRACT": 0x6D,
    "DECIMAL": 0x6E,
    "DIVIDE": 0x6F,
    "F1": 0x70,
    "F2": 0x71,
    "F3": 0x72,
    "F4": 0x73,
    "F5": 0x74,
    "F6": 0x75,
    "F7": 0x76,
    "F8": 0x77,
    "F9": 0x78,
    "F10": 0x79,
    "F11": 0x7A,
    "F12": 0x7B,
    "F13": 0x7C,
    "F14": 0x7D,
    "F15": 0x7E,
    "F16": 0x7F,
    "F17": 0x80,
    "F18": 0x81,
    "F19": 0x82,
    "F20": 0x83,
    "F21": 0x84,
    "F22": 0x85,
    "F23": 0x86,
    "F24": 0x87,
    "NUMLOCK": 0x90,
    "SCROLL": 0x91,
    "LSHIFT": 0xA0,
    "RSHIFT": 0xA1,
    "LCONTROL": 0xA2,
    "RCONTROL": 0xA3,
    "LMENU": 0xA4,
    "RMENU": 0xA5,
    "BROWSER_BACK": 0xA6,
    "BROWSER_FORWARD": 0xA7,
    "BROWSER_REFRESH": 0xA8,
    "BROWSER_STOP": 0xA9,
    "BROWSER_SEARCH": 0xAA,
    "BROWSER_FAVORITES": 0xAB,
    "BROWSER_HOME": 0xAC,
    "VOLUME_MUTE": 0xAD,
    "VOLUME_DOWN": 0xAE,
    "VOLUME_UP": 0xAF,
    "MEDIA_NEXT_TRACK": 0xB0,
    "MEDIA_PREV_TRACK": 0xB1,
    "MEDIA_STOP": 0xB2,
    "MEDIA_PLAY_PAUSE": 0xB3,
    "LAUNCH_MAIL": 0xB4,
    "LAUNCH_MEDIA_SELECT": 0xB5,
    "LAUNCH_APP1": 0xB6,
    "LAUNCH_APP2": 0xB7,
    "OEM_1": 0xBA,
    "OEM_PLUS": 0xBB,
    "OEM_COMMA": 0xBC,
    "OEM_MINUS": 0xBD,
    "OEM_PERIOD": 0xBE,
    "OEM_2": 0xBF,
    "OEM_3": 0xC0,
    "OEM_4": 0xDB,
    "OEM_5": 0xDC,
    "OEM_6": 0xDD,
    "OEM_7": 0xDE,
    "OEM_8": 0xDF,
    "OEM_102": 0xE2,
    "PROCESSKEY": 0xE5,
    "PACKET": 0xE7,
    "ATTN": 0xF6,
    "CRSEL": 0xF7,
    "EXSEL": 0xF8,
    "EREOF": 0xF9,
    "PLAY": 0xFA,
    "ZOOM": 0xFB,
    "NONAME": 0xFC,
    "PA1": 0xFD,
    "OEM_CLEAR": 0xFE,
}

# Define ULONG_PTR for compatibility if not present in ctypes.wintypes
if not hasattr(wintypes, "ULONG_PTR"):
    import sys

    if sys.maxsize > 2**32:
        wintypes.ULONG_PTR = ctypes.c_uint64
    else:
        wintypes.ULONG_PTR = ctypes.c_uint32


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
    ]


class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]

    _anonymous_ = ("_input",)
    _fields_ = [("type", wintypes.DWORD), ("_input", _INPUT)]


def _send_input(input_obj: Union[INPUT, list[INPUT]]) -> None:
    if isinstance(input_obj, list):
        n_inputs = len(input_obj)
        arr_type = INPUT * n_inputs
        arr = arr_type(*input_obj)
        user32.SendInput(n_inputs, ctypes.byref(arr), ctypes.sizeof(INPUT))
    else:
        user32.SendInput(1, ctypes.byref(input_obj), ctypes.sizeof(input_obj))


_CLSID_CUIAutomation = com.guid("{FF48DBA4-60EF-4201-AA87-54103EEF594E}")
_IID_IUIAutomation = com.guid("{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}")

_IID_InvokePattern = com.guid("{FB377FBE-8EA6-46D5-9C73-6499642D3059}")
_IID_ValuePattern = com.guid("{A94CD8B1-0844-4CD6-9D2D-640537AB39E9}")
_IID_ExpandCollapsePattern = com.guid("{619BE086-1F4E-4EE4-BAFA-210128738730}")
_IID_TogglePattern = com.guid("{94CF8058-9B8D-4AB9-8BFD-4CD0A33C8C70}")
_IID_ScrollPattern = com.guid("{88F4D42A-E881-4573-A11C-05A5B44E9170}")
_IID_SelectionItemPattern = com.guid("{A8EFA66A-0FDA-421A-9194-38021F3578EA}")
_IID_ScrollItemPattern = com.guid("{B488300F-D015-4F19-9C29-BB595E3645EF}")

_PAT_INVOKE = 10000
_PAT_VALUE = 10002
_PAT_SCROLL = 10004
_PAT_EXPAND_COLLAPSE = 10005
_PAT_SELECTION_ITEM = 10010
_PAT_TOGGLE = 10015
_PAT_SCROLL_ITEM = 10017

_PAT_IID: dict[int, com.GUID] = {
    _PAT_INVOKE: _IID_InvokePattern,
    _PAT_VALUE: _IID_ValuePattern,
    _PAT_SCROLL: _IID_ScrollPattern,
    _PAT_EXPAND_COLLAPSE: _IID_ExpandCollapsePattern,
    _PAT_SELECTION_ITEM: _IID_SelectionItemPattern,
    _PAT_TOGGLE: _IID_TogglePattern,
    _PAT_SCROLL_ITEM: _IID_ScrollItemPattern,
}

_UIA_GET_ROOT = 5
_UIA_CONTROL_VIEW_WALKER = 14
_EL_SET_FOCUS = 3
_EL_GET_RUNTIME_ID = 4
_EL_GET_PATTERN_AS = 14
_EL_CONTROL_TYPE = 21
_EL_NAME = 23
_EL_IS_ENABLED = 28
_EL_AUTOMATION_ID = 29
_EL_HAS_KEYBOARD_FOCUS = 26
_EL_IS_KEYBOARD_FOCUSABLE = 27
_EL_IS_OFFSCREEN = 38
_EL_BOUNDING_RECT = 43
_TW_FIRST_CHILD = 4
_TW_NEXT_SIBLING = 6
_INV_INVOKE = 3
_VAL_SET_VALUE = 3
_VAL_CURRENT_VALUE = 4
_VAL_IS_READONLY = 5
_TOG_TOGGLE = 3
_TOG_CURRENT_STATE = 4
_EC_EXPAND = 3
_EC_COLLAPSE = 4
_EC_CURRENT_STATE = 5
_SCR_SCROLL = 3
_SEL_SELECT = 3
_SCI_SCROLL_INTO_VIEW = 3

_CONTROL_TYPES: dict[int, str] = {
    50000: "button",
    50001: "calendar",
    50002: "checkbox",
    50003: "combobox",
    50004: "edit",
    50005: "hyperlink",
    50006: "image",
    50007: "listitem",
    50008: "list",
    50009: "menu",
    50010: "menubar",
    50011: "menuitem",
    50012: "progressbar",
    50013: "radiobutton",
    50014: "scrollbar",
    50015: "slider",
    50016: "spinner",
    50017: "statusbar",
    50018: "tab",
    50019: "tabitem",
    50020: "text",
    50021: "toolbar",
    50022: "tooltip",
    50023: "tree",
    50024: "treeitem",
    50025: "custom",
    50026: "group",
    50027: "thumb",
    50028: "datagrid",
    50029: "dataitem",
    50030: "document",
    50031: "splitbutton",
    50032: "window",
    50033: "pane",
    50034: "header",
    50035: "headeritem",
    50036: "table",
    50037: "titlebar",
    50038: "separator",
}

_MAX_NODES = 500

_AX_OUT_VP = (ctypes.POINTER(ctypes.c_void_p),)
_AX_OUT_INT = (ctypes.POINTER(ctypes.c_int),)
_AX_OUT_BOOL = (ctypes.POINTER(wintypes.BOOL),)
_AX_OUT_RECT = (ctypes.POINTER(wintypes.RECT),)


def _el_control_type(el: int) -> Optional[int]:
    v = ctypes.c_int()
    hr = com.vc(el, _EL_CONTROL_TYPE, ctypes.c_long, _AX_OUT_INT, ctypes.byref(v))
    return v.value if hr >= 0 else None


def _el_is_offscreen(el: int) -> bool:
    v = wintypes.BOOL()
    hr = com.vc(el, _EL_IS_OFFSCREEN, ctypes.c_long, _AX_OUT_BOOL, ctypes.byref(v))
    return bool(v.value) if hr >= 0 else True


def _el_name(el: int) -> Optional[str]:
    b = ctypes.c_void_p()
    hr = com.vc(el, _EL_NAME, ctypes.c_long, _AX_OUT_VP, ctypes.byref(b))
    return com.bstr_val(b.value) if hr >= 0 else None


def _el_automation_id(el: int) -> Optional[str]:
    b = ctypes.c_void_p()
    hr = com.vc(el, _EL_AUTOMATION_ID, ctypes.c_long, _AX_OUT_VP, ctypes.byref(b))
    return com.bstr_val(b.value) if hr >= 0 else None


def _el_is_enabled(el: int) -> bool:
    v = wintypes.BOOL()
    hr = com.vc(el, _EL_IS_ENABLED, ctypes.c_long, _AX_OUT_BOOL, ctypes.byref(v))
    return bool(v.value) if hr >= 0 else True


def _el_has_keyboard_focus(el: int) -> bool:
    v = wintypes.BOOL()
    hr = com.vc(
        el, _EL_HAS_KEYBOARD_FOCUS, ctypes.c_long, _AX_OUT_BOOL, ctypes.byref(v)
    )
    return bool(v.value) if hr >= 0 else False


def _el_is_keyboard_focusable(el: int) -> bool:
    v = wintypes.BOOL()
    hr = com.vc(
        el, _EL_IS_KEYBOARD_FOCUSABLE, ctypes.c_long, _AX_OUT_BOOL, ctypes.byref(v)
    )
    return bool(v.value) if hr >= 0 else False


def _el_bounding_rect(el: int) -> Optional[wintypes.RECT]:
    r = wintypes.RECT()
    hr = com.vc(el, _EL_BOUNDING_RECT, ctypes.c_long, _AX_OUT_RECT, ctypes.byref(r))
    return r if hr >= 0 else None


def _el_runtime_id(el: int) -> Optional[list[int]]:
    psa = ctypes.c_void_p()
    hr = com.vc(el, _EL_GET_RUNTIME_ID, ctypes.c_long, _AX_OUT_VP, ctypes.byref(psa))
    return com.sa_ints(psa.value) if hr >= 0 else None


def _el_set_focus(el: int) -> int:
    return com.vc(el, _EL_SET_FOCUS, ctypes.c_long, ())


def _el_pattern(el: int, pat_id: int) -> Optional[int]:
    iid = _PAT_IID.get(pat_id)
    if iid is None:
        return None
    out = ctypes.c_void_p()
    hr = com.vc(
        el,
        _EL_GET_PATTERN_AS,
        ctypes.c_long,
        (ctypes.c_int, ctypes.POINTER(com.GUID), ctypes.POINTER(ctypes.c_void_p)),
        pat_id,
        ctypes.byref(iid),
        ctypes.byref(out),
    )
    return out.value if hr >= 0 and out.value else None


def _el_center(el: int) -> tuple[int, int]:
    r = _el_bounding_rect(el)
    if r is None:
        raise OperatorRuntimeException("Cannot get bounding rectangle")
    return (r.left + r.right) // 2, (r.top + r.bottom) // 2


_TW_ARGS = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p))


def _tw_first_child(walker: int, el: int) -> Optional[int]:
    out = ctypes.c_void_p()
    hr = com.vc(walker, _TW_FIRST_CHILD, ctypes.c_long, _TW_ARGS, el, ctypes.byref(out))
    return out.value if hr >= 0 and out.value else None


def _tw_next_sibling(walker: int, el: int) -> Optional[int]:
    out = ctypes.c_void_p()
    hr = com.vc(
        walker, _TW_NEXT_SIBLING, ctypes.c_long, _TW_ARGS, el, ctypes.byref(out)
    )
    return out.value if hr >= 0 and out.value else None


def _uia_root(uia: int) -> int:
    out = ctypes.c_void_p()
    hr = com.vc(uia, _UIA_GET_ROOT, ctypes.c_long, _AX_OUT_VP, ctypes.byref(out))
    if hr < 0 or not out.value:
        raise OperatorRuntimeException("Failed to get UIA root element")
    return out.value


def _uia_walker(uia: int) -> int:
    out = ctypes.c_void_p()
    hr = com.vc(
        uia, _UIA_CONTROL_VIEW_WALKER, ctypes.c_long, _AX_OUT_VP, ctypes.byref(out)
    )
    if hr < 0 or not out.value:
        raise OperatorRuntimeException("Failed to get UIA ControlViewWalker")
    return out.value


def _pat_invoke(p: int) -> None:
    com.vc(p, _INV_INVOKE, ctypes.c_long, ())


def _pat_toggle(p: int) -> None:
    com.vc(p, _TOG_TOGGLE, ctypes.c_long, ())


def _pat_toggle_state(p: int) -> int:
    v = ctypes.c_int()
    com.vc(p, _TOG_CURRENT_STATE, ctypes.c_long, _AX_OUT_INT, ctypes.byref(v))
    return v.value


def _pat_select(p: int) -> None:
    com.vc(p, _SEL_SELECT, ctypes.c_long, ())


def _pat_expand(p: int) -> None:
    com.vc(p, _EC_EXPAND, ctypes.c_long, ())


def _pat_collapse(p: int) -> None:
    com.vc(p, _EC_COLLAPSE, ctypes.c_long, ())


def _pat_ec_state(p: int) -> int:
    v = ctypes.c_int()
    com.vc(p, _EC_CURRENT_STATE, ctypes.c_long, _AX_OUT_INT, ctypes.byref(v))
    return v.value


def _pat_value_get(p: int) -> Optional[str]:
    b = ctypes.c_void_p()
    hr = com.vc(p, _VAL_CURRENT_VALUE, ctypes.c_long, _AX_OUT_VP, ctypes.byref(b))
    return com.bstr_val(b.value) if hr >= 0 else None


def _pat_value_is_readonly(p: int) -> bool:
    v = wintypes.BOOL()
    hr = com.vc(p, _VAL_IS_READONLY, ctypes.c_long, _AX_OUT_BOOL, ctypes.byref(v))
    return bool(v.value) if hr >= 0 else False


def _pat_value_set(p: int, text: str) -> int:
    b = com.bstr_alloc(text)
    try:
        return com.vc(p, _VAL_SET_VALUE, ctypes.c_long, (ctypes.c_void_p,), b)
    finally:
        com.bstr_free(b)


def _pat_scroll(p: int, h_amount: int, v_amount: int) -> int:
    return com.vc(
        p, _SCR_SCROLL, ctypes.c_long, (ctypes.c_int, ctypes.c_int), h_amount, v_amount
    )


def _pat_scroll_into_view(p: int) -> int:
    return com.vc(p, _SCI_SCROLL_INTO_VIEW, ctypes.c_long, ())


class Win32CuaOperator(BaseCuaOperator):
    async def move(self, x: int, y: int) -> None:
        screen_width = user32.GetSystemMetrics(0)
        screen_height = user32.GetSystemMetrics(1)
        abs_x = int(x * 65535 / screen_width)
        abs_y = int(y * 65535 / screen_height)
        inp = INPUT(
            type=INPUT_MOUSE,
            mi=MOUSEINPUT(
                dx=abs_x,
                dy=abs_y,
                mouseData=0,
                dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
                time=0,
                dwExtraInfo=0,
            ),
        )
        _send_input(inp)

    async def click(
        self, button: str = "left", x: Optional[int] = None, y: Optional[int] = None
    ) -> None:
        if x is not None and y is not None:
            await self.move(x, y)
            await asyncio.sleep(SLEEP_INTERVAL)
        if button == "left":
            down_flag = MOUSEEVENTF_LEFTDOWN
            up_flag = MOUSEEVENTF_LEFTUP
            mouseData = 0
        elif button == "right":
            down_flag = MOUSEEVENTF_RIGHTDOWN
            up_flag = MOUSEEVENTF_RIGHTUP
            mouseData = 0
        elif button == "middle" or button == "wheel":
            down_flag = 0x0020
            up_flag = 0x0040
            mouseData = 0
        elif button == "back":
            down_flag = 0x0080
            up_flag = 0x0100
            mouseData = 0x0001
        elif button == "forward":
            down_flag = 0x0080
            up_flag = 0x0100
            mouseData = 0x0002
        else:
            raise OperatorRuntimeException(
                "button must be 'left', 'right', 'middle', 'wheel', 'back', or 'forward'"
            )
        down = INPUT(type=INPUT_MOUSE, mi=MOUSEINPUT(0, 0, mouseData, down_flag, 0, 0))
        up = INPUT(type=INPUT_MOUSE, mi=MOUSEINPUT(0, 0, mouseData, up_flag, 0, 0))
        _send_input([down, up])

    async def double_click(
        self, x: Optional[int] = None, y: Optional[int] = None
    ) -> None:
        await self.click(button="left", x=x, y=y)
        await asyncio.sleep(SLEEP_INTERVAL)
        await self.click(button="left", x=x, y=y)

    async def drag(self, path: list[tuple[int, int]]) -> None:
        if not path or len(path) < 2:
            raise OperatorRuntimeException("Path must contain at least two points")
        await self.move(*path[0])
        await asyncio.sleep(SLEEP_INTERVAL)
        down = INPUT(
            type=INPUT_MOUSE, mi=MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTDOWN, 0, 0)
        )
        _send_input(down)
        await asyncio.sleep(SLEEP_INTERVAL)
        for point in path[1:-1]:
            await self.move(*point)
            await asyncio.sleep(SLEEP_INTERVAL)
        await self.move(*path[-1])
        await asyncio.sleep(SLEEP_INTERVAL)
        up = INPUT(type=INPUT_MOUSE, mi=MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTUP, 0, 0))
        _send_input(up)

    async def key_press(self, keys: list[str]) -> None:
        vk_codes = []
        for key in keys:
            key_upper = key.upper()
            if key_upper in VK_MAP:
                vk_codes.append(VK_MAP[key_upper])
            elif len(key) == 1:
                vk = user32.VkKeyScanW(ord(key)) & 0xFF
                if vk != 0xFF:
                    vk_codes.append(vk)
            else:
                raise OperatorRuntimeException(f"Unknown key: {key}")

        downs = [
            INPUT(
                type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=0),
            )
            for vk in vk_codes
        ]
        ups = [
            INPUT(
                type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(
                    wVk=vk, wScan=0, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0
                ),
            )
            for vk in reversed(vk_codes)
        ]
        _send_input(downs)
        await asyncio.sleep(SLEEP_INTERVAL)
        _send_input(ups)

    async def type_text(self, text: str) -> None:
        for char in text:
            await asyncio.sleep(SLEEP_INTERVAL)
            vk_scan = user32.VkKeyScanW(ord(char))
            vk = vk_scan & 0xFF
            shift = (vk_scan >> 8) & 0xFF
            if vk == 0xFF or ord(char) > 0x7F:
                # Encode to UTF-16-LE to correctly handle surrogate pairs (chars > U+FFFF)
                utf16 = char.encode("utf-16-le")
                for i in range(0, len(utf16), 2):
                    scan_code = int.from_bytes(utf16[i : i + 2], "little")
                    down = INPUT(
                        type=INPUT_KEYBOARD,
                        ki=KEYBDINPUT(
                            wVk=0,
                            wScan=scan_code,
                            dwFlags=KEYEVENTF_UNICODE,
                            time=0,
                            dwExtraInfo=0,
                        ),
                    )
                    up = INPUT(
                        type=INPUT_KEYBOARD,
                        ki=KEYBDINPUT(
                            wVk=0,
                            wScan=scan_code,
                            dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                            time=0,
                            dwExtraInfo=0,
                        ),
                    )
                    _send_input(down)
                    await asyncio.sleep(SLEEP_INTERVAL)
                    _send_input(up)
            else:
                downs = []
                ups = []
                if shift & 1:
                    downs.append(
                        INPUT(
                            type=INPUT_KEYBOARD,
                            ki=KEYBDINPUT(
                                wVk=VK_SHIFT, wScan=0, dwFlags=0, time=0, dwExtraInfo=0
                            ),
                        )
                    )
                downs.append(
                    INPUT(
                        type=INPUT_KEYBOARD,
                        ki=KEYBDINPUT(
                            wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=0
                        ),
                    )
                )
                ups.append(
                    INPUT(
                        type=INPUT_KEYBOARD,
                        ki=KEYBDINPUT(
                            wVk=vk,
                            wScan=0,
                            dwFlags=KEYEVENTF_KEYUP,
                            time=0,
                            dwExtraInfo=0,
                        ),
                    )
                )
                if shift & 1:
                    ups.append(
                        INPUT(
                            type=INPUT_KEYBOARD,
                            ki=KEYBDINPUT(
                                wVk=VK_SHIFT,
                                wScan=0,
                                dwFlags=KEYEVENTF_KEYUP,
                                time=0,
                                dwExtraInfo=0,
                            ),
                        )
                    )

                _send_input(downs)
                await asyncio.sleep(SLEEP_INTERVAL)
                _send_input(ups)

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
        if scroll_y != 0:
            inp = INPUT(
                type=INPUT_MOUSE, mi=MOUSEINPUT(0, 0, scroll_y, MOUSEEVENTF_WHEEL, 0, 0)
            )
            _send_input(inp)
        if scroll_x != 0:
            inp = INPUT(
                type=INPUT_MOUSE,
                mi=MOUSEINPUT(0, 0, scroll_x, MOUSEEVENTF_HWHEEL, 0, 0),
            )
            _send_input(inp)

    async def screenshot(self) -> str:
        img = ImageGrab.grab()
        screen_width = user32.GetSystemMetrics(0)
        screen_height = user32.GetSystemMetrics(1)
        if img.size != (screen_width, screen_height):
            img = img.resize((screen_width, screen_height), Image.LANCZOS)
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
        return img_str

    async def dimensions(self) -> tuple[int, int]:
        screen_width = user32.GetSystemMetrics(0)
        screen_height = user32.GetSystemMetrics(1)
        return screen_width, screen_height

    async def monitors(self) -> list[MonitorMetadata]:
        try:
            shcore = ctypes.windll.shcore
        except OSError:
            shcore = None

        result: list[MonitorMetadata] = []

        MONITORINFOF_PRIMARY = 0x00000001

        class MONITORINFOEX(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("rcMonitor", wintypes.RECT),
                ("rcWork", wintypes.RECT),
                ("dwFlags", wintypes.DWORD),
                ("szDevice", wintypes.WCHAR * 32),
            ]

        MONITORENUMPROC = ctypes.WINFUNCTYPE(
            wintypes.BOOL,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.POINTER(wintypes.RECT),
            wintypes.LPARAM,
        )

        def _callback(hmonitor, _hdc, _lprect, _lparam):
            info = MONITORINFOEX()
            info.cbSize = ctypes.sizeof(MONITORINFOEX)
            user32.GetMonitorInfoW(hmonitor, ctypes.byref(info))

            width = info.rcMonitor.right - info.rcMonitor.left
            height = info.rcMonitor.bottom - info.rcMonitor.top
            is_primary = bool(info.dwFlags & MONITORINFOF_PRIMARY)

            scale = 1.0
            if shcore:
                dpi_x = ctypes.c_uint()
                dpi_y = ctypes.c_uint()
                hr = shcore.GetDpiForMonitor(
                    hmonitor, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y)
                )
                if hr == 0:
                    scale = dpi_x.value / 96.0

            result.append(
                MonitorMetadata(
                    id=hmonitor,
                    width=width,
                    height=height,
                    scale=scale,
                    primary=is_primary,
                )
            )
            return True

        enum_proc = MONITORENUMPROC(_callback)
        user32.EnumDisplayMonitors(None, None, enum_proc, 0)
        return result

    async def wait(self) -> None:
        await asyncio.sleep(1)

    def _ax_init(self) -> None:
        if hasattr(self, "_ax_walker"):
            return
        self._uia: int = com.co_create_instance(
            _CLSID_CUIAutomation, _IID_IUIAutomation
        )
        self._ax_walker: int = _uia_walker(self._uia)
        self._ax_cache: dict[str, int] = {}
        self._ax_node_count: int = 0

    def _ax_element(self, node_id: str) -> int:
        el = self._ax_cache.get(node_id)
        if el is None:
            raise OperatorRuntimeException(
                f"Node '{node_id}' not found. Call axtree() first."
            )
        return el

    def _ax_clear_cache(self) -> None:
        for p in self._ax_cache.values():
            try:
                com.release(p)
            except (ValueError, OSError):
                pass
        self._ax_cache.clear()

    @staticmethod
    def _ax_rid_str(el: int) -> Optional[str]:
        rid = _el_runtime_id(el)
        if rid:
            return ".".join(str(v) for v in rid)
        return None

    def _ax_walk(self, el: int, depth: int, max_depth: int) -> Optional[AxNode]:
        if depth > max_depth or self._ax_node_count >= _MAX_NODES or not el:
            return None

        try:
            ct = _el_control_type(el)
        except ValueError:
            return None
        if ct is None:
            return None
        if _el_is_offscreen(el):
            return None

        node_id = self._ax_rid_str(el)
        if node_id is None:
            return None
        self._ax_node_count += 1
        self._ax_cache[node_id] = el

        role = _CONTROL_TYPES.get(ct, "unknown")
        name = _el_name(el) or ""

        states: list[AxNodeState] = [AxNodeState.VISIBLE]
        if _el_is_enabled(el):
            states.append(AxNodeState.ENABLED)
        if _el_has_keyboard_focus(el):
            states.append(AxNodeState.FOCUSED)
        if _el_is_keyboard_focusable(el):
            states.append(AxNodeState.FOCUSABLE)

        sp = _el_pattern(el, _PAT_SELECTION_ITEM)
        if sp:
            v = wintypes.BOOL()
            hr = com.vc(sp, 6, ctypes.c_long, _AX_OUT_BOOL, ctypes.byref(v))
            if hr >= 0 and v.value:
                states.append(AxNodeState.SELECTED)
            com.release(sp)

        ecp = _el_pattern(el, _PAT_EXPAND_COLLAPSE)
        if ecp:
            if _pat_ec_state(ecp) == 1:  # ExpandCollapseState.Expanded
                states.append(AxNodeState.EXPANDED)
            com.release(ecp)

        tp = _el_pattern(el, _PAT_TOGGLE)
        if tp:
            if _pat_toggle_state(tp) == 1:  # ToggleState.On
                states.append(AxNodeState.CHECKED)
            com.release(tp)

        vp = _el_pattern(el, _PAT_VALUE)
        if vp:
            if _pat_value_is_readonly(vp):
                states.append(AxNodeState.READONLY)
            com.release(vp)

        bounds: Optional[AxNodeBounds] = None
        rect = _el_bounding_rect(el)
        if rect:
            bounds = AxNodeBounds(
                x=rect.left,
                y=rect.top,
                w=rect.right - rect.left,
                h=rect.bottom - rect.top,
            )

        children: list[AxNode] = []
        child = _tw_first_child(self._ax_walker, el)
        while child:
            child_node = self._ax_walk(child, depth + 1, max_depth)
            if child_node is not None:
                children.append(child_node)
            nxt = _tw_next_sibling(self._ax_walker, child)
            if child_node is None:
                com.release(child)
            child = nxt

        return AxNode(
            id=node_id,
            role=role,
            name=name,
            bounds=bounds,
            states=states,
            children=children,
        )

    async def axtree(
        self, root_node_id: Optional[str] = None, max_depth: int = 8
    ) -> AxNode:
        self._ax_init()
        if root_node_id:
            root = self._ax_element(root_node_id)
            com.vc(root, 1, ctypes.c_ulong, ())  # AddRef before cache clear
        else:
            root = None

        self._ax_clear_cache()
        self._ax_node_count = 0

        if root is None:
            root = _uia_root(self._uia)

        tree = self._ax_walk(root, 0, max_depth)
        if tree is None:
            com.release(root)
            return AxNode(id="", role="empty")
        return tree

    async def ax_click(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)

        p = _el_pattern(el, _PAT_INVOKE)
        if p:
            _pat_invoke(p)
            com.release(p)
            return

        p = _el_pattern(el, _PAT_TOGGLE)
        if p:
            _pat_toggle(p)
            com.release(p)
            return

        p = _el_pattern(el, _PAT_SELECTION_ITEM)
        if p:
            _pat_select(p)
            com.release(p)
            return

        cx, cy = _el_center(el)
        await self.click(button="left", x=cx, y=cy)

    async def ax_double_click(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        cx, cy = _el_center(el)
        await self.double_click(x=cx, y=cy)

    async def ax_type(self, text: str, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        vp = _el_pattern(el, _PAT_VALUE)
        if vp:
            hr = _pat_value_set(vp, text)
            com.release(vp)
            if hr >= 0:
                return
        _el_set_focus(el)
        await asyncio.sleep(0.05)
        await self.type_text(text)

    async def ax_scroll(
        self,
        scroll_x: int,
        scroll_y: int,
        node_id: str,
    ) -> None:
        self._ax_init()
        el = self._ax_element(node_id)

        sp = _el_pattern(el, _PAT_SCROLL)
        if sp:
            h = self._ax_scroll_amount(scroll_x)
            v = self._ax_scroll_amount(scroll_y)
            hr = _pat_scroll(sp, h, v)
            com.release(sp)
            if hr >= 0:
                return

        sip = _el_pattern(el, _PAT_SCROLL_ITEM)
        if sip:
            _pat_scroll_into_view(sip)
            com.release(sip)

        cx, cy = _el_center(el)
        await self.scroll(scroll_x, scroll_y, x=cx, y=cy)

    async def ax_focus(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        hr = _el_set_focus(el)
        if hr < 0:
            raise OperatorRuntimeException(
                f"Cannot focus node '{node_id}': HRESULT 0x{hr & 0xFFFFFFFF:08X}"
            )

    async def ax_expand(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        p = _el_pattern(el, _PAT_EXPAND_COLLAPSE)
        if p is None:
            raise OperatorRuntimeException(
                f"Node '{node_id}' does not support expand/collapse."
            )
        _pat_expand(p)
        com.release(p)

    async def ax_collapse(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        p = _el_pattern(el, _PAT_EXPAND_COLLAPSE)
        if p is None:
            raise OperatorRuntimeException(
                f"Node '{node_id}' does not support expand/collapse."
            )
        _pat_collapse(p)
        com.release(p)

    async def ax_select(self, node_id: str) -> None:
        self._ax_init()
        el = self._ax_element(node_id)
        p = _el_pattern(el, _PAT_SELECTION_ITEM)
        if p is None:
            raise OperatorRuntimeException(
                f"Node '{node_id}' does not support selection."
            )
        _pat_select(p)
        com.release(p)

    @staticmethod
    def _ax_scroll_amount(value: int) -> int:
        if value == 0:
            return 2  # NoAmount
        if abs(value) >= 120:
            return 3 if value > 0 else 0  # Large
        return 4 if value > 0 else 1  # Small
