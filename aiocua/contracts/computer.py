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


class AxAction(StrEnum):
    CLICK = "click"
    FOCUS = "focus"
    TYPE = "type"
    SCROLL = "scroll"
    EXPAND = "expand"
    COLLAPSE = "collapse"
    SELECT = "select"


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
    secondary_role: Optional[str] = None
    name: str = ""
    bounds: Optional[AxNodeBounds] = None
    states: list[AxNodeState] = field(default_factory=list)
    children: list["AxNode"] = field(default_factory=list)
    description: str = ""
    value: Optional[str] = None
    allowed_actions: list[AxAction] = field(default_factory=list)


@dataclass
class MonitorMetadata:
    id: int
    width: int
    height: int
    scale: float
    primary: bool
