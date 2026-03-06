from dataclasses import dataclass, field
from enum import StrEnum
from typing import Optional


class AxNodeState(StrEnum):
    ENABLED = "enabled"
    FOCUSED = "focused"
    FOCUSABLE = "focusable"
    SELECTED = "selected"
    EXPANDED = "expanded"
    CHECKED = "checked"
    VISIBLE = "visible"
    READONLY = "readonly"


@dataclass
class AxNodeBounds:
    x: int
    y: int
    w: int
    h: int


@dataclass
class AxNode:
    id: str
    role: str
    name: str = ""
    bounds: Optional[AxNodeBounds] = None
    states: list[AxNodeState] = field(default_factory=list)
    children: list["AxNode"] = field(default_factory=list)


@dataclass
class MonitorMetadata:
    id: int
    width: int
    height: int
    scale: float
    primary: bool
