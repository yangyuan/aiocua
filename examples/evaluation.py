import asyncio
import sys, os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from aiocua import CuaOperator


async def main():
    operator = CuaOperator()

    x, y = await operator.dimensions()
    await asyncio.sleep(1)
    await operator.move(500, 300)
    await asyncio.sleep(1)
    for button in ["left", "right", "wheel", "back", "forward"]:
        await operator.click(button=button)
        await asyncio.sleep(1)
    await operator.double_click()
    await asyncio.sleep(1)
    await operator.key_press(["A"])
    await asyncio.sleep(1)
    await operator.key_press(["SHIFT", "A"])
    await asyncio.sleep(1)
    await operator.type_text("Hello World!")
    await asyncio.sleep(1)
    await operator.type_text("中文")
    await asyncio.sleep(1)
    await operator.drag([(300, 300), (600, 600)])
    await asyncio.sleep(1)
    await operator.scroll(scroll_x=0, scroll_y=410)
    await asyncio.sleep(1)
    await operator.scroll(scroll_x=410, scroll_y=0)
    await asyncio.sleep(1)
    screenshot_base64 = await operator.screenshot()
    await operator.wait()


if __name__ == "__main__":
    asyncio.run(main())
