from dataclasses import dataclass


@dataclass
class MonitorMetadata:
    id: int
    width: int
    height: int
    scale: float
    primary: bool
