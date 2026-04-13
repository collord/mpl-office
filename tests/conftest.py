"""Shared pytest fixtures."""
from __future__ import annotations

from pathlib import Path

import pytest

ARTIFACTS = Path(__file__).parent / "_artifacts"


@pytest.fixture(scope="session", autouse=True)
def _ensure_artifacts_dir():
    ARTIFACTS.mkdir(exist_ok=True)
    return ARTIFACTS


@pytest.fixture
def artifacts() -> Path:
    return ARTIFACTS
