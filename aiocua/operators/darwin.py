import asyncio
import ctypes
import ctypes.util
import base64
from io import BytesIO
from typing import Optional
from PIL import Image
from aiocua.contracts.computer import MonitorMetadata
from aiocua.contracts.error import OperatorRuntimeException
from aiocua.operators.base import BaseCuaOperator

core = ctypes.cdll.LoadLibrary(ctypes.util.find_library("CoreGraphics"))


class CGPoint(ctypes.Structure):
    _fields_ = [("x", ctypes.c_double), ("y", ctypes.c_double)]


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


def _get_mouse_pos() -> tuple[float, float]:
    event = core.CGEventCreate(None)
    pt = core.CGEventGetLocation(event)
    core.CFRelease(event)
    return pt.x, pt.y


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

    async def axtree(self, root_node_id=None, max_depth=8):
        raise NotImplementedError("axtree is not implemented on macOS")

    async def ax_click(self, node_id):
        raise NotImplementedError("ax_click is not implemented on macOS")

    async def ax_double_click(self, node_id):
        raise NotImplementedError("ax_double_click is not implemented on macOS")

    async def ax_type(self, text, node_id):
        raise NotImplementedError("ax_type is not implemented on macOS")

    async def ax_scroll(self, scroll_x, scroll_y, node_id):
        raise NotImplementedError("ax_scroll is not implemented on macOS")

    async def ax_focus(self, node_id):
        raise NotImplementedError("ax_focus is not implemented on macOS")

    async def ax_expand(self, node_id):
        raise NotImplementedError("ax_expand is not implemented on macOS")

    async def ax_collapse(self, node_id):
        raise NotImplementedError("ax_collapse is not implemented on macOS")

    async def ax_select(self, node_id):
        raise NotImplementedError("ax_select is not implemented on macOS")
