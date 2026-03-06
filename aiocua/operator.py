import sys
from typing import Optional

from aiocua.contracts.computer import AxNode, MonitorMetadata
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
            raise OperatorRuntimeException(
                "x and y must both be provided or both be None"
            )
        return await self._operator.click(button, x, y)

    async def double_click(
        self, x: Optional[int] = None, y: Optional[int] = None
    ) -> None:
        if (x is None) != (y is None):
            raise OperatorRuntimeException(
                "x and y must both be provided or both be None"
            )
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
            raise OperatorRuntimeException(
                "x and y must both be provided or both be None"
            )
        return await self._operator.scroll(scroll_x, scroll_y, x, y)

    async def screenshot(self) -> str:
        return await self._operator.screenshot()

    async def dimensions(self) -> tuple[int, int]:
        return await self._operator.dimensions()

    async def monitors(self) -> list[MonitorMetadata]:
        return await self._operator.monitors()

    async def wait(self) -> None:
        return await self._operator.wait()

    async def axtree(
        self, root_node_id: Optional[str] = None, max_depth: int = 8
    ) -> AxNode:
        return await self._operator.axtree(root_node_id, max_depth)

    async def ax_click(self, node_id: str) -> None:
        return await self._operator.ax_click(node_id)

    async def ax_double_click(self, node_id: str) -> None:
        return await self._operator.ax_double_click(node_id)

    async def ax_type(self, text: str, node_id: str) -> None:
        return await self._operator.ax_type(text, node_id)

    async def ax_scroll(
        self,
        scroll_x: int,
        scroll_y: int,
        node_id: str,
    ) -> None:
        return await self._operator.ax_scroll(scroll_x, scroll_y, node_id)

    async def ax_focus(self, node_id: str) -> None:
        return await self._operator.ax_focus(node_id)

    async def ax_expand(self, node_id: str) -> None:
        return await self._operator.ax_expand(node_id)

    async def ax_collapse(self, node_id: str) -> None:
        return await self._operator.ax_collapse(node_id)

    async def ax_select(self, node_id: str) -> None:
        return await self._operator.ax_select(node_id)
