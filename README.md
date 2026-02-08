# aiocua

A Simple Python CUA (Computer Use Agent) Solution

## Status

This project is in **Technical Preview**.
- APIs are subject to change.
- Not ready for production use.

## Supported Platforms

- Windows
- macOS

## Installation

```bash
pip install aiocua
```

## Quick Start

```python
import asyncio
from aiocua import CuaOperator

async def main():
    operator = CuaOperator()

    # Get screen dimensions
    width, height = await operator.dimensions()

    # Take a screenshot (returns base64-encoded string)
    screenshot_base64 = await operator.screenshot()

    # Get monitor info
    monitors = await operator.monitors()
    for m in monitors:
        print(f"Monitor {m.id}: {m.width}x{m.height} (scale={m.scale}, primary={m.primary})")

if __name__ == "__main__":
    asyncio.run(main())
```

## API Reference

All methods on `CuaOperator` are async.

### Mouse

| Method | Description |
| --- | --- |
| `move(x, y)` | Move the cursor to the given coordinates. |
| `click(button="left", x=None, y=None)` | Click a mouse button. Supported buttons: `left`, `right`, `wheel`, `back`, `forward`. If `x`/`y` are provided, moves first. |
| `double_click(x=None, y=None)` | Double-click at the current or given position. |
| `drag(path)` | Drag along a list of `(x, y)` coordinates. |
| `scroll(scroll_x, scroll_y, x=None, y=None)` | Scroll by the given amounts. Optionally move to `x`/`y` first. |

### Keyboard

| Method | Description |
| --- | --- |
| `key_press(keys)` | Press one or more keys simultaneously (e.g. `["CTRL", "C"]`). Keys are automatically remapped per platform. |
| `type_text(text)` | Type a string of text, including Unicode characters. |

### Screen

| Method | Description |
| --- | --- |
| `screenshot()` | Capture the screen and return it as a base64-encoded PNG string. |
| `dimensions()` | Return the primary monitor resolution as `(width, height)`. |
| `monitors()` | Return a list of `MonitorMetadata` for each connected display. |

### Other

| Method | Description |
| --- | --- |
| `wait()` | Wait for any pending operations to complete. |

## Testing

```bash
pip install aiocua[test]
pytest
```
