"""Shared Inputs/Outputs path helpers for every SA skill.

All skills under ``SA\\Skills\\`` share one drop folder for source files
(``SA\\Inputs``) and one destination root for generated reports
(``SA\\Outputs\\<skill-name>\\``), so a skill never needs a hardcoded
output path and a user never has to remember a different input location
per skill.

Usage from a skill script (siblings of this file's parent folder)::

    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
    from sa_paths import output_dir_for, latest_input

    out_dir = output_dir_for("my-skill-name")
    src = args.input or latest_input("*.xlsx", "*.csv")
"""
from __future__ import annotations

from pathlib import Path

SA_ROOT = Path(__file__).resolve().parent.parent.parent
INPUTS_DIR = SA_ROOT / "Inputs"
OUTPUTS_DIR = SA_ROOT / "Outputs"


def output_dir_for(skill_name: str) -> Path:
    """Return (creating if needed) the shared output folder for a skill."""
    d = OUTPUTS_DIR / skill_name
    d.mkdir(parents=True, exist_ok=True)
    return d


def latest_input(*patterns: str) -> Path | None:
    """Return the most recently modified file in the shared Inputs folder
    matching any of the given glob patterns (e.g. "*.xlsx", "*.csv"), or
    None if nothing matches. Used only when a skill's own --input-style
    flag isn't explicitly provided."""
    candidates: list[Path] = []
    if INPUTS_DIR.exists():
        for pat in patterns:
            candidates.extend(p for p in INPUTS_DIR.glob(pat) if p.is_file())
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)
