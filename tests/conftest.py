"""测试公共夹具：在临时目录生成一份测试数据，并隔离用户别名字典。"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from generate_test_data import generate  # noqa: E402


@pytest.fixture(scope="session")
def data_dir(tmp_path_factory) -> Path:
    return generate(tmp_path_factory.mktemp("data") / "test_data")


@pytest.fixture(autouse=True)
def isolated_home(tmp_path, monkeypatch) -> Path:
    """每个测试使用独立的用户数据目录，避免学习到的别名互相影响。"""
    home = tmp_path / "home"
    monkeypatch.setenv("SHEETMERGER_HOME", str(home))
    return home
