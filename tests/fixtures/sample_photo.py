"""A minimal valid 1x1 PNG, hand-written so no imaging library is a
dependency (the spec limits Python dependencies to pywin32 and pytest)."""
from __future__ import annotations

from pathlib import Path

_PNG_BYTES = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
    "53de0000000c4944415408d763f8ffff3f0005fe02fea739669d0000000049"
    "454e44ae426082"
)


def write_sample_photo(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(_PNG_BYTES)
    return path
