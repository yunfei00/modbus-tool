"""
Modbus 客户端封装：与 pymodbus 同步客户端交互，供 UI 调用。

兼容 pymodbus 3.x 中 unit/slave/device_id 命名差异，以及写单寄存器方法名差异。
"""

from __future__ import annotations

import inspect
from typing import Any

from pymodbus.client import ModbusSerialClient, ModbusTcpClient
from pymodbus.exceptions import (
    ConnectionException,
    ModbusException,
    ModbusIOException,
)

# 客户端类型别名（TCP / 串口共用 Mixin 接口）
_Client = ModbusTcpClient | ModbusSerialClient


def _unit_kwargs(unit_id: int) -> dict[str, Any]:
    """
    根据当前 pymodbus 版本构造「从站地址」关键字参数。
    3.13+ 常用 device_id；更早版本常用 slave 或 unit。
    """
    sig = inspect.signature(ModbusTcpClient.read_holding_registers)
    if "device_id" in sig.parameters:
        return {"device_id": unit_id}
    if "slave" in sig.parameters:
        return {"slave": unit_id}
    if "unit" in sig.parameters:
        return {"unit": unit_id}
    return {"device_id": unit_id}


def _format_modbus_error(response: Any) -> str:
    """将 pymodbus 响应对象转为可读错误说明。"""
    parts: list[str] = [str(response)]
    if getattr(response, "exception_code", None) is not None:
        parts.append(f"异常码={response.exception_code}")
    if getattr(response, "function_code", None) is not None:
        parts.append(f"功能码={response.function_code}")
    return " | ".join(parts)


class ModbusClient:
    """薄封装：管理连接生命周期，并提供读写寄存器方法。"""

    def __init__(self) -> None:
        self._client: _Client | None = None

    def _require_client(self) -> _Client:
        if self._client is None:
            raise RuntimeError("客户端未创建")
        return self._client

    def _write_single_register(self, address: int, value: int, unit_kw: dict[str, Any]) -> Any:
        """调用写单个保持寄存器（适配 write_register / write_single_register）。"""
        client = self._require_client()
        if hasattr(client, "write_register"):
            return client.write_register(address, value, **unit_kw)
        if hasattr(client, "write_single_register"):
            return client.write_single_register(address, value, **unit_kw)
        raise RuntimeError("当前 pymodbus 版本不支持写单个寄存器接口")

    def _call_pymodbus_write_registers(
        self, address: int, values: list[int], unit_kw: dict[str, Any]
    ) -> Any:
        """调用 pymodbus 的 write_registers，避免与对外 API 同名。"""
        client = self._require_client()
        return client.write_registers(address, values, **unit_kw)

    # ------------------------------------------------------------------ 连接
    def connect_tcp(self, host: str, port: int) -> None:
        """
        建立 Modbus TCP 连接。
        :raises ConnectionException: 库内部连接异常
        :raises ModbusIOException: IO 层错误
        """
        self.close()
        client = ModbusTcpClient(host, port=port, timeout=3)
        try:
            if not client.connect():
                client.close()
                raise ConnectionException(f"无法连接到 {host}:{port}")
        except Exception:
            client.close()
            raise
        self._client = client

    def connect_rtu(
        self,
        port: str,
        baudrate: int,
        bytesize: int,
        parity: str,
        stopbits: int | float,
    ) -> None:
        """
        建立 Modbus RTU（串口）连接。
        :raises ConnectionException: 无法打开串口或握手失败
        """
        self.close()
        # framer 默认 RTU，无需显式传入 Framer 枚举以保持多版本兼容
        client = ModbusSerialClient(
            port,
            baudrate=baudrate,
            bytesize=bytesize,
            parity=parity,
            stopbits=stopbits,
            timeout=3,
        )
        try:
            if not client.connect():
                client.close()
                raise ConnectionException(f"无法打开串口或连接失败: {port}")
        except Exception:
            client.close()
            raise
        self._client = client

    def close(self) -> None:
        """断开并释放客户端。"""
        if self._client is not None:
            try:
                self._client.close()
            finally:
                self._client = None

    def is_connected(self) -> bool:
        """是否已与设备保持连接。"""
        if self._client is None:
            return False
        try:
            return bool(self._client.connected)
        except Exception:
            return bool(getattr(self._client, "is_socket_open", lambda: False)())

    # ------------------------------------------------------------------ 读
    def read_holding_registers(self, unit_id: int, address: int, count: int) -> list[int]:
        """
        功能码 03：读保持寄存器。
        :return: 寄存器值列表（每个元素 0~65535）
        """
        client = self._require_client()
        unit_kw = _unit_kwargs(unit_id)
        try:
            resp = client.read_holding_registers(address, count=count, **unit_kw)
        except (ModbusIOException, ConnectionException, ModbusException) as exc:
            raise RuntimeError(f"通讯异常: {exc}") from exc
        if resp.isError():
            raise RuntimeError(f"Modbus 异常响应: {_format_modbus_error(resp)}")
        return list(resp.registers)

    def read_input_registers(self, unit_id: int, address: int, count: int) -> list[int]:
        """
        功能码 04：读输入寄存器。
        """
        client = self._require_client()
        unit_kw = _unit_kwargs(unit_id)
        try:
            resp = client.read_input_registers(address, count=count, **unit_kw)
        except (ModbusIOException, ConnectionException, ModbusException) as exc:
            raise RuntimeError(f"通讯异常: {exc}") from exc
        if resp.isError():
            raise RuntimeError(f"Modbus 异常响应: {_format_modbus_error(resp)}")
        return list(resp.registers)

    # ------------------------------------------------------------------ 写
    def write_single_register(self, unit_id: int, address: int, value: int) -> None:
        """功能码 06：写单个保持寄存器。"""
        client = self._require_client()
        unit_kw = _unit_kwargs(unit_id)
        try:
            resp = self._write_single_register(address, value, unit_kw)
        except (ModbusIOException, ConnectionException, ModbusException) as exc:
            raise RuntimeError(f"通讯异常: {exc}") from exc
        if resp.isError():
            raise RuntimeError(f"Modbus 异常响应: {_format_modbus_error(resp)}")

    def write_multiple_registers(self, unit_id: int, address: int, values: list[int]) -> None:
        """功能码 16：写多个保持寄存器。"""
        client = self._require_client()
        unit_kw = _unit_kwargs(unit_id)
        try:
            resp = self._call_pymodbus_write_registers(address, values, unit_kw)
        except (ModbusIOException, ConnectionException, ModbusException) as exc:
            raise RuntimeError(f"通讯异常: {exc}") from exc
        if resp.isError():
            raise RuntimeError(f"Modbus 异常响应: {_format_modbus_error(resp)}")
