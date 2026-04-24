"""
连接与界面选项的 JSON 配置保存/加载。

默认文件：项目根目录下 config/modbus_tool_config.json
（以包上级目录为「项目根」，即包含 main.py 的目录）。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict


def project_root() -> Path:
    """modbus_tool 包的上级目录（仓库根）。"""
    return Path(__file__).resolve().parents[2]


def default_config_path() -> Path:
    """默认配置文件路径；若 config 目录不存在则创建。"""
    cfg_dir = project_root() / "config"
    try:
        cfg_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return cfg_dir / "modbus_tool_config.json"


def save_config(path: Path, data: Dict[str, Any]) -> None:
    """将字典写入 JSON 文件（UTF-8）。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(data, ensure_ascii=False, indent=2)
    path.write_text(text, encoding="utf-8")


def load_config(path: Path) -> Dict[str, Any]:
    """读取 JSON 配置文件；文件不存在或损坏时抛出异常。"""
    raw = path.read_text(encoding="utf-8")
    obj = json.loads(raw)
    if not isinstance(obj, dict):
        raise ValueError("配置文件根节点必须是 JSON 对象")
    return obj
