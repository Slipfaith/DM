# config.py
from dataclasses import dataclass

@dataclass
class Config:
    header_color: int = 65535  # Yellow
    dry_run: bool = False