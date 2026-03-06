from abc import ABC, abstractmethod
from typing import Optional

from aiocua.contracts.computer import AxNode, MonitorMetadata


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

    @abstractmethod
    async def axtree(
        self, root_node_id: Optional[str] = None, max_depth: int = 8
    ) -> AxNode: ...

    @abstractmethod
    async def ax_click(self, node_id: str) -> None: ...

    @abstractmethod
    async def ax_double_click(self, node_id: str) -> None: ...

    @abstractmethod
    async def ax_type(self, text: str, node_id: str) -> None: ...

    @abstractmethod
    async def ax_scroll(
        self,
        scroll_x: int,
        scroll_y: int,
        node_id: str,
    ) -> None: ...

    @abstractmethod
    async def ax_focus(self, node_id: str) -> None: ...

    @abstractmethod
    async def ax_expand(self, node_id: str) -> None: ...

    @abstractmethod
    async def ax_collapse(self, node_id: str) -> None: ...

    @abstractmethod
    async def ax_select(self, node_id: str) -> None: ...
