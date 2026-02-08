from abc import ABC, abstractmethod
from typing import Optional

from aiocua.contracts.computer import MonitorMetadata


class BaseCuaOperator(ABC):
    @abstractmethod
    async def move(self, x: int, y: int) -> None: ...

    @abstractmethod
    async def click(
        self, button: str = "left", x: Optional[int] = None, y: Optional[int] = None
    ) -> None: ...

    @abstractmethod
    async def double_click(
        self, x: Optional[int] = None, y: Optional[int] = None
    ) -> None: ...

    @abstractmethod
    async def drag(self, path: list[tuple[int, int]]) -> None: ...

    @abstractmethod
    async def key_press(self, keys: list[str]) -> None: ...

    @abstractmethod
    async def type_text(self, text: str) -> None: ...

    @abstractmethod
    async def scroll(
        self,
        scroll_x: int,
        scroll_y: int,
        x: Optional[int] = None,
        y: Optional[int] = None,
    ) -> None: ...

    @abstractmethod
    async def screenshot(self) -> str: ...

    @abstractmethod
    async def dimensions(self) -> tuple[int, int]: ...

    @abstractmethod
    async def monitors(self) -> list[MonitorMetadata]: ...

    @abstractmethod
    async def wait(self) -> None: ...
