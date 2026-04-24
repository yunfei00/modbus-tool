"""
串口枚举：基于 pyserial 的 list_ports，供 UI 刷新下拉列表。
"""

from __future__ import annotations

from typing import List

try:
    from serial.tools import list_ports as _list_ports
except ImportError:
    _list_ports = None  # type: ignore[misc, assignment]


def list_ports_available() -> bool:
    """本机是否可使用 serial.tools.list_ports 枚举串口。"""
    return _list_ports is not None


def list_serial_devices() -> List[str]:
    """
    返回当前系统可用串口设备名列表（如 Windows 下 COM3）。
    若 pyserial 不可用或枚举失败，返回空列表。
    """
    if _list_ports is None:
        return []
    try:
        return [p.device for p in _list_ports.comports()]
    except Exception:  # noqa: BLE001 — 枚举失败时返回空，由上层记日志
        return []
