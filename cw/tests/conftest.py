"""Shared test configuration. The repository root is the parent of this folder."""
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def repo_root():
    return REPO_ROOT
