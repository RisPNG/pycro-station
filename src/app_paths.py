from __future__ import annotations

from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SRC_DIR.parent

PYCROS_DIR = PROJECT_ROOT / "pycros"
REMOTE_PYCROS_DIR = PROJECT_ROOT / "remote_pycros"
REQUIREMENTS_TXT = PROJECT_ROOT / "requirements.txt"


def project_path(*parts: str) -> Path:
    return PROJECT_ROOT.joinpath(*parts)
