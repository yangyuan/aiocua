import sys
from typing import Optional

from aiocua.contracts.computer import MonitorMetadata
from aiocua.contracts.error import OperatorRuntimeException
from aiocua.operators.base import BaseCuaOperator
from aiocua.services.keyboard import KeyboardService


class CuaOperator:
    def __init__(self) -> None:
        self._keyboard_service = KeyboardService(sys.platform)

        if sys.platform == "win32":
            from aiocua.operators.win32 import Win32CuaOperator

            self._operator: BaseCuaOperator = Win32CuaOperator()
        elif sys.platform == "darwin":
            from aiocua.operators.darwin import DarwinCuaOperator

            self._operator = DarwinCuaOperator()
        else:
            raise NotImplementedError(f"Unsupported platform: {sys.platform}")

    async def move(self, x: int, y: int) -> None:
        return await self._operator.move(x, y)

    async def click(
        self, button: str = "left", x: Optional[int] = None, y: Optional[int] = None
    ) -> None:
        if (x is None) != (y is None):
            raise OperatorRuntimeException("x and y must both be provided or both be None")
        return await self._operator.click(button, x, y)

    async def double_click(
        self, x: Optional[int] = None, y: Optional[int] = None
    ) -> None:
        if (x is None) != (y is None):
            raise OperatorRuntimeException("x and y must both be provided or both be None")
        return await self._operator.double_click(x, y)

    async def drag(self, path: list[tuple[int, int]]) -> None:
        return await self._operator.drag(path)

    async def key_press(self, keys: list[str]) -> None:
        remapped = self._keyboard_service.remap_keys(keys)
        return await self._operator.key_press(remapped)

    async def type_text(self, text: str) -> None:
        return await self._operator.type_text(text)

    async def scroll(
        self,
        scroll_x: int,
        scroll_y: int,
        x: Optional[int] = None,
        y: Optional[int] = None,
    ) -> None:
        if (x is None) != (y is None):
            raise OperatorRuntimeException("x and y must both be provided or both be None")
        return await self._operator.scroll(scroll_x, scroll_y, x, y)

    async def screenshot(self) -> str:
        return await self._operator.screenshot()

    async def dimensions(self) -> tuple[int, int]:
        return await self._operator.dimensions()

    async def monitors(self) -> list[MonitorMetadata]:
        return await self._operator.monitors()

    async def wait(self) -> None:
        return await self._operator.wait()

