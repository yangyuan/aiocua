_DARWIN_KEY_REMAP: dict[str, str] = {
    "/": "/",
    "\\": "\\",
    "ALT": "ALT",
    "ARROWDOWN": "DOWN",
    "ARROWLEFT": "LEFT",
    "ARROWRIGHT": "RIGHT",
    "ARROWUP": "UP",
    "BACKSPACE": "DELETE",
    "CAPSLOCK": "CAPSLOCK",
    "CMD": "CMD",
    "CTRL": "CTRL",
    "DEL": "DELETE",
    "DELETE": "DELETE",
    "END": "END",
    "ENTER": "RETURN",
    "ESC": "ESC",
    "HOME": "HOME",
    "INSERT": "INSERT",
    "OPTION": "ALT",
    "PAGEDOWN": "PAGEDOWN",
    "PAGEUP": "PAGEUP",
    "SHIFT": "SHIFT",
    "SPACE": "SPACE",
    "SUPER": "CMD",
    "TAB": "TAB",
    "WIN": "CMD",
}

_WIN32_KEY_REMAP: dict[str, str] = {
    "/": "/",
    "\\": "\\",
    "ALT": "MENU",
    "ARROWDOWN": "DOWN",
    "ARROWLEFT": "LEFT",
    "ARROWRIGHT": "RIGHT",
    "ARROWUP": "UP",
    "BACKSPACE": "BACK",
    "CAPSLOCK": "CAPITAL",
    "CMD": "LWIN",
    "CTRL": "CONTROL",
    "DEL": "DELETE",
    "DELETE": "DELETE",
    "END": "END",
    "ENTER": "RETURN",
    "ESC": "ESCAPE",
    "HOME": "HOME",
    "INSERT": "INSERT",
    "OPTION": "MENU",
    "PAGEDOWN": "NEXT",
    "PAGEUP": "PRIOR",
    "SHIFT": "SHIFT",
    "SPACE": "SPACE",
    "SUPER": "LWIN",
    "TAB": "TAB",
    "WIN": "LWIN",
}


class KeyboardService:
    def __init__(self, platform: str) -> None:
        self._platform = platform
        if platform == "darwin":
            self._key_remap = _DARWIN_KEY_REMAP
        elif platform == "win32":
            self._key_remap = _WIN32_KEY_REMAP
        else:
            raise NotImplementedError(f"Unsupported platform: {platform}")

    def remap_keys(self, keys: list[str]) -> list[str]:
        return [self._key_remap.get(k.upper(), k) for k in keys]
