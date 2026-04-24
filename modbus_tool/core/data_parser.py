"""
寄存器原始值到数据类型的解析（uint16 / int16 / uint32 / int32 / float32）。

字序：
- 「AB CD」：高 16 位字在前（先读到的寄存器为高字）。
- 「CD AB」：低 16 位字在前（字交换）。
"""

from __future__ import annotations

import struct
from typing import List


def _combine_u32(reg_hi: int, reg_lo: int, word_order: str) -> int:
    """将两个 uint16 寄存器组合为 32 位无符号整数。"""
    a = reg_hi & 0xFFFF
    b = reg_lo & 0xFFFF
    if word_order == "AB CD":
        return (a << 16) | b
    if word_order == "CD AB":
        return (b << 16) | a
    raise ValueError(f"不支持的字序: {word_order!r}")


def _u32_to_int32(u: int) -> int:
    u &= 0xFFFFFFFF
    if u > 0x7FFFFFFF:
        return u - 0x100000000
    return u


def _u32_to_float32(u: int) -> float:
    """将 32 位无符号模式按 IEEE754 big-endian 解释为 float32。"""
    u &= 0xFFFFFFFF
    b = struct.pack(">I", u)
    return struct.unpack(">f", b)[0]


def build_parsed_column(
    registers: List[int],
    dtype: str,
    word_order: str,
) -> List[str]:
    """
    为每个寄存器行生成「解析值」列文本。

    - uint16 / int16：每行一个解析值。
    - uint32 / int32 / float32：每两个字解析一个值，解析结果显示在高字所在行，
      配对中的第二行解析列为空字符串；末尾单独一个寄存器时解析为空。
    """
    n = len(registers)
    out: List[str] = [""] * n
    dtype = dtype.lower().strip()

    if dtype == "uint16":
        for i in range(n):
            out[i] = str(registers[i] & 0xFFFF)
        return out

    if dtype == "int16":
        for i in range(n):
            v = registers[i] & 0xFFFF
            if v >= 0x8000:
                v -= 0x10000
            out[i] = str(v)
        return out

    if dtype in ("uint32", "int32", "float32"):
        i = 0
        while i < n:
            if i + 1 < n:
                r0, r1 = registers[i], registers[i + 1]
                u = _combine_u32(r0, r1, word_order)
                if dtype == "uint32":
                    out[i] = str(u & 0xFFFFFFFF)
                elif dtype == "int32":
                    out[i] = str(_u32_to_int32(u))
                else:
                    try:
                        out[i] = str(_u32_to_float32(u))
                    except (struct.error, OverflowError):
                        out[i] = ""
                out[i + 1] = ""
                i += 2
            else:
                out[i] = ""
                i += 1
        return out

    # 未知类型：不抛异常；解析列为空
    return out
