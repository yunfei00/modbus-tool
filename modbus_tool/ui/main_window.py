"""
主窗口：连接配置、读写操作、结果表格与日志。
所有 pymodbus 调用均通过 ModbusClient，本文件不直接依赖 pymodbus。

布局：左侧固定宽度控制区；右侧自适应结果表 + 通信日志。
"""

from __future__ import annotations

import traceback
from datetime import datetime
from typing import Iterable

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

try:
    from serial.tools import list_ports
except ImportError:
    list_ports = None  # type: ignore[misc, assignment]

from modbus_tool.core.modbus_client import ModbusClient


def _log_time() -> str:
    """日志行内时间戳 [HH:mm:ss]。"""
    return datetime.now().strftime("%H:%M:%S")


def _fmt_reg_hex(val: int) -> str:
    """十六进制四位大写，如 0x000A。"""
    return f"0x{(val & 0xFFFF):04X}"


def _fmt_reg_bin(val: int) -> str:
    """16 位二进制前缀 0b。"""
    return "0b" + format(val & 0xFFFF, "016b")


class MainWindow(QWidget):
    """Modbus Tool V1 主界面。"""

    FC_OPTIONS = (
        ("03", "03 Read Holding Registers"),
        ("04", "04 Read Input Registers"),
        ("06", "06 Write Single Register"),
        ("16", "16 Write Multiple Registers"),
    )

    # 左侧控制条固定宽度（像素），落在 360~420 区间内
    LEFT_PANEL_WIDTH = 400

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Modbus Tool V1")
        self.resize(1280, 760)

        self._client = ModbusClient()

        self._build_ui()
        self._wire_signals()
        self._update_fc_dependent_widgets()
        self._set_connection_status(False)
        self.append_log("INFO", "工具启动完成。")
        # 启动时枚举串口不写刷新日志，避免先于「工具启动」刷屏
        self._refresh_serial_ports(log_info=False)

    # ------------------------------------------------------------------ UI 构建
    def _build_ui(self) -> None:
        root = QHBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(10, 10, 10, 10)

        # ======================== 左侧控制区 ========================
        left_panel = QWidget()
        left_panel.setFixedWidth(self.LEFT_PANEL_WIDTH)
        left_col = QVBoxLayout(left_panel)
        left_col.setSpacing(10)

        # --- 连接配置
        conn_group = QGroupBox("连接配置")
        conn_layout = QVBoxLayout(conn_group)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("通讯类型:"))
        self.combo_conn = QComboBox()
        self.combo_conn.addItems(["TCP", "RTU"])
        row_type.addWidget(self.combo_conn)
        row_type.addStretch()
        conn_layout.addLayout(row_type)

        # TCP / RTU 参数分页，切换类型时只显示对应参数
        self.stack_params = QStackedWidget()
        tcp_page = QWidget()
        tcp_form = QFormLayout(tcp_page)
        self.edit_host = QLineEdit("127.0.0.1")
        self.spin_port = QSpinBox()
        self.spin_port.setRange(1, 65535)
        self.spin_port.setValue(502)
        tcp_form.addRow("IP:", self.edit_host)
        tcp_form.addRow("端口:", self.spin_port)
        self.stack_params.addWidget(tcp_page)

        rtu_page = QWidget()
        rtu_grid = QGridLayout(rtu_page)
        self.combo_serial = QComboBox()
        self.combo_serial.setEditable(True)
        self.combo_serial.setMinimumContentsLength(16)
        self.btn_refresh_ports = QPushButton("刷新串口")
        rtu_grid.addWidget(QLabel("串口:"), 0, 0)
        rtu_grid.addWidget(self.combo_serial, 0, 1)
        rtu_grid.addWidget(self.btn_refresh_ports, 0, 2)

        self.spin_baud = QSpinBox()
        self.spin_baud.setRange(1200, 921600)
        self.spin_baud.setValue(9600)
        self.spin_bytesize = QComboBox()
        self.spin_bytesize.addItems(["7", "8"])
        self.spin_bytesize.setCurrentText("8")
        self.combo_parity = QComboBox()
        self.combo_parity.addItems(["N", "E", "O"])
        self.combo_stopbits = QComboBox()
        self.combo_stopbits.addItems(["1", "1.5", "2"])
        self.combo_stopbits.setCurrentText("1")

        rtu_grid.addWidget(QLabel("波特率:"), 1, 0)
        rtu_grid.addWidget(self.spin_baud, 1, 1)
        rtu_grid.addWidget(QLabel("数据位:"), 2, 0)
        rtu_grid.addWidget(self.spin_bytesize, 2, 1)
        rtu_grid.addWidget(QLabel("校验位:"), 3, 0)
        rtu_grid.addWidget(self.combo_parity, 3, 1)
        rtu_grid.addWidget(QLabel("停止位:"), 4, 0)
        rtu_grid.addWidget(self.combo_stopbits, 4, 1)

        self.stack_params.addWidget(rtu_page)
        conn_layout.addWidget(self.stack_params)

        row_unit = QHBoxLayout()
        row_unit.addWidget(QLabel("从站地址 (Unit ID):"))
        self.spin_unit = QSpinBox()
        self.spin_unit.setRange(1, 247)
        self.spin_unit.setValue(1)
        row_unit.addWidget(self.spin_unit)
        row_unit.addStretch()
        conn_layout.addLayout(row_unit)

        row_btns = QHBoxLayout()
        self.btn_connect = QPushButton("连接")
        self.btn_disconnect = QPushButton("断开")
        self.btn_disconnect.setEnabled(False)
        row_btns.addWidget(self.btn_connect)
        row_btns.addWidget(self.btn_disconnect)
        row_btns.addStretch()
        conn_layout.addLayout(row_btns)

        left_col.addWidget(conn_group)

        # --- 操作
        op_group = QGroupBox("操作")
        op_layout = QFormLayout(op_group)

        self.combo_fc = QComboBox()
        for code, label in self.FC_OPTIONS:
            self.combo_fc.addItem(label, userData=code)
        op_layout.addRow("功能码:", self.combo_fc)

        self.spin_addr = QSpinBox()
        self.spin_addr.setRange(0, 65535)
        self.spin_addr.setValue(0)
        op_layout.addRow("起始地址:", self.spin_addr)

        self.spin_count = QSpinBox()
        self.spin_count.setRange(1, 125)
        self.spin_count.setValue(10)
        op_layout.addRow("数量:", self.spin_count)

        self.edit_values = QLineEdit()
        op_layout.addRow("写入值:", self.edit_values)

        self.btn_exec = QPushButton("执行")
        op_layout.addRow(self.btn_exec)

        left_col.addWidget(op_group)

        # 弹性空白，将连接状态压在左侧底部
        left_col.addStretch(1)

        # --- 连接状态（左侧底部，圆点 + 文案）
        status_group = QGroupBox("连接状态")
        status_layout = QHBoxLayout(status_group)
        self.lbl_status_icon = QLabel("●")
        self.lbl_status_icon.setFixedWidth(22)
        self.lbl_status_icon.setAlignment(Qt.AlignCenter)
        self.lbl_status_text = QLabel("状态：未连接")
        status_layout.addWidget(self.lbl_status_icon)
        status_layout.addWidget(self.lbl_status_text, stretch=1)
        left_col.addWidget(status_group)

        root.addWidget(left_panel, stretch=0)

        # ======================== 右侧显示区 ========================
        right_panel = QWidget()
        right_col = QVBoxLayout(right_panel)
        right_col.setSpacing(8)

        splitter = QSplitter(Qt.Vertical)

        res_group = QGroupBox("结果")
        res_layout = QVBoxLayout(res_group)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["地址", "十进制", "十六进制", "二进制"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        res_layout.addWidget(self.table)
        splitter.addWidget(res_group)

        log_group = QGroupBox("通信日志")
        log_layout = QVBoxLayout(log_group)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(200)
        _mono = QFont("Consolas", 10)
        if not _mono.exactMatch():
            _mono = QFont("Courier New", 10)
        self.log.setFont(_mono)
        log_layout.addWidget(self.log)
        splitter.addWidget(log_group)

        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 3)
        right_col.addWidget(splitter)
        root.addWidget(right_panel, stretch=1)

    def _wire_signals(self) -> None:
        self.combo_conn.currentIndexChanged.connect(self._on_conn_type_changed)
        self.btn_refresh_ports.clicked.connect(self._refresh_serial_ports)
        self.btn_connect.clicked.connect(self._on_connect)
        self.btn_disconnect.clicked.connect(self._on_disconnect)
        self.btn_exec.clicked.connect(self._on_execute)
        self.combo_fc.currentIndexChanged.connect(self._update_fc_dependent_widgets)

    # ------------------------------------------------------------------ 日志 / 提示
    def append_log(self, level: str, message: str) -> None:
        """
        追加一行日志。
        格式：[HH:mm:ss] [LEVEL] 内容
        LEVEL: INFO / OK / WARN / ERROR / TX / RX
        """
        level = level.upper()
        self.log.append(f"[{_log_time()}] [{level}] {message}")

    def _show_error(self, title: str, message: str) -> None:
        """错误：弹窗 + ERROR 级别日志。"""
        self.append_log("ERROR", f"{title}: {message}")
        QMessageBox.warning(self, title, message)

    def _show_warning(self, title: str, message: str) -> None:
        """警告：弹窗 + WARN 级别日志。"""
        self.append_log("WARN", f"{title}: {message}")
        QMessageBox.warning(self, title, message)

    def _set_exec_busy(self, busy: bool) -> None:
        """执行按钮忙碌态（执行中禁用并改文案）。"""
        self.btn_exec.setEnabled(not busy)
        self.btn_exec.setText("执行中..." if busy else "执行")

    def _set_connection_status(self, connected: bool) -> None:
        """更新左下角连接状态指示（红/绿）。"""
        if connected:
            self.lbl_status_icon.setStyleSheet("color: #43A047; font-size: 16px;")
            self.lbl_status_text.setText("状态：已连接")
            self.lbl_status_text.setStyleSheet("color: #2E7D32; font-weight: bold;")
        else:
            self.lbl_status_icon.setStyleSheet("color: #E53935; font-size: 16px;")
            self.lbl_status_text.setText("状态：未连接")
            self.lbl_status_text.setStyleSheet("color: #C62828; font-weight: bold;")

    # ------------------------------------------------------------------ 串口列表
    def _refresh_serial_ports(self, log_info: bool = True) -> None:
        """枚举本机串口并填充下拉框（需要 pyserial）。"""
        current = self.combo_serial.currentText().strip()
        self.combo_serial.clear()
        if list_ports is None:
            # 缺少枚举能力时始终提示一次，避免静默失败
            self.append_log(
                "WARN",
                "无法导入 serial.tools.list_ports，请手动输入串口名。",
            )
            if current:
                self.combo_serial.addItem(current)
                self.combo_serial.setEditText(current)
            return
        ports = list(list_ports.comports())
        for p in ports:
            desc = f"{p.device} — {p.description}" if p.description else p.device
            self.combo_serial.addItem(desc, userData=p.device)
        idx = self.combo_serial.findData(current)
        if idx >= 0:
            self.combo_serial.setCurrentIndex(idx)
        elif current:
            self.combo_serial.setEditText(current)
        elif self.combo_serial.count():
            self.combo_serial.setCurrentIndex(0)
        if log_info:
            self.append_log("INFO", f"串口列表已刷新，共 {len(ports)} 个端口。")

    def _current_serial_device(self) -> str:
        """从下拉框解析真实串口设备名（Windows: COMx）。"""
        i = self.combo_serial.currentIndex()
        if i >= 0:
            data = self.combo_serial.itemData(i)
            if data:
                return str(data)
        return self.combo_serial.currentText().strip().split("—")[0].strip()

    # ------------------------------------------------------------------ 连接类型切换
    def _on_conn_type_changed(self, _index: int) -> None:
        is_tcp = self.combo_conn.currentText() == "TCP"
        self.stack_params.setCurrentIndex(0 if is_tcp else 1)

    # ------------------------------------------------------------------ 功能码相关控件状态
    def _current_fc(self) -> str:
        return str(self.combo_fc.currentData())

    def _update_fc_dependent_widgets(self) -> None:
        fc = self._current_fc()
        is_read = fc in ("03", "04")
        is_single = fc == "06"
        is_multi = fc == "16"

        self.spin_count.setEnabled(is_read or is_multi)
        self.edit_values.setEnabled(is_single or is_multi)

        if is_read:
            self.edit_values.clear()
            self.edit_values.setPlaceholderText("读操作无需填写写入值")
        elif is_single:
            self.edit_values.setPlaceholderText("请输入单个寄存器值，例如 123")
        else:
            self.edit_values.setPlaceholderText("请输入多个寄存器值，例如 1,2,3,4")

    def _set_connected_ui(self, connected: bool) -> None:
        self.btn_connect.setEnabled(not connected)
        self.btn_disconnect.setEnabled(connected)
        self.combo_conn.setEnabled(not connected)
        self.stack_params.setEnabled(not connected)
        self.spin_unit.setEnabled(not connected)
        self.btn_refresh_ports.setEnabled(not connected)
        self._set_connection_status(connected)

    # ------------------------------------------------------------------ 连接 / 断开
    def _on_connect(self) -> None:
        if self._client.is_connected():
            self._show_warning("提示", "已经处于连接状态。")
            return

        try:
            if self.combo_conn.currentText() == "TCP":
                host = self.edit_host.text().strip()
                if not host:
                    self._show_error("参数错误", "IP 地址不能为空。")
                    return
                port = int(self.spin_port.value())
                self.append_log("INFO", f"正在连接 TCP {host}:{port}")
                self._client.connect_tcp(host, port)
                self.append_log("OK", f"TCP 连接成功：{host}:{port}")
            else:
                port_name = self._current_serial_device()
                if not port_name:
                    self._show_error("参数错误", "请选择或填写串口号。")
                    return
                baud = int(self.spin_baud.value())
                bytesize = int(self.spin_bytesize.currentText())
                parity = self.combo_parity.currentText()
                stop = float(self.combo_stopbits.currentText())
                self.append_log(
                    "INFO",
                    f"正在连接 RTU 串口 {port_name}，"
                    f"{baud}/{bytesize}{parity}{stop}",
                )
                self._client.connect_rtu(port_name, baud, bytesize, parity, stop)
                self.append_log("OK", f"RTU 连接成功：{port_name}")
        except Exception as exc:  # noqa: BLE001
            self._client.close()
            msg = self._format_user_exception(exc, context="连接失败")
            self._show_error("连接失败", msg)
            self._set_connection_status(False)
            return

        self._set_connected_ui(True)

    def _on_disconnect(self) -> None:
        self._client.close()
        self._set_connected_ui(False)
        self.append_log("INFO", "已断开连接。")

    # ------------------------------------------------------------------ 执行读写
    def _parse_write_values_single(self, text: str) -> int:
        s = text.strip()
        if not s:
            raise ValueError("写入值不能为空。")
        try:
            v = int(s, 0)
        except ValueError as exc:
            raise ValueError("写入值必须是整数（可带 0x 前缀）。") from exc
        if not 0 <= v <= 65535:
            raise ValueError("寄存器值必须在 0~65535 范围内。")
        return v

    def _parse_write_values_multi(self, text: str) -> list[int]:
        s = text.strip()
        if not s:
            raise ValueError("写入值不能为空。")
        parts = [p.strip() for p in s.split(",") if p.strip()]
        if not parts:
            raise ValueError("请使用英文逗号分隔多个整数，例如: 1,2,3")
        values: list[int] = []
        for p in parts:
            try:
                v = int(p, 0)
            except ValueError as exc:
                raise ValueError(f"无法解析的数值: {p!r}") from exc
            if not 0 <= v <= 65535:
                raise ValueError(f"数值超出 0~65535: {p!r}")
            values.append(v)
        return values

    def _format_user_exception(self, exc: BaseException, context: str = "") -> str:
        """将异常转换为用户可读说明（含无响应 / Modbus 异常）。"""
        text = str(exc).strip() or type(exc).__name__
        lower = text.lower()
        if "modbus 异常响应" in text:
            return text
        if isinstance(exc, TimeoutError) or "timeout" in lower or "timed out" in lower:
            return "设备无响应或请求超时。请检查线路、从站地址与波特率等参数。"
        if "no response" in lower or "modbusio" in type(exc).__name__.lower():
            return f"设备无响应: {text}"
        prefix = f"{context} — " if context else ""
        return f"{prefix}{text}"

    def _clear_result_table(self) -> None:
        """清空结果表（写成功后使用）。"""
        self.table.setRowCount(0)

    def _fill_table(self, start_addr: int, registers: Iterable[int]) -> None:
        """用读取结果填充表格（四列）。"""
        self.table.setRowCount(0)
        for i, val in enumerate(registers):
            row = self.table.rowCount()
            self.table.insertRow(row)
            addr = start_addr + i
            self.table.setItem(row, 0, QTableWidgetItem(str(addr)))
            self.table.setItem(row, 1, QTableWidgetItem(str(val)))
            self.table.setItem(row, 2, QTableWidgetItem(_fmt_reg_hex(val)))
            self.table.setItem(row, 3, QTableWidgetItem(_fmt_reg_bin(val)))

    def _on_execute(self) -> None:
        if not self._client.is_connected():
            self._show_error("未连接", "请先连接设备后再执行操作。")
            return

        self._set_exec_busy(True)
        try:
            unit_id = int(self.spin_unit.value())
            address = int(self.spin_addr.value())
            count = int(self.spin_count.value())
            fc = self._current_fc()
            raw_values = self.edit_values.text()

            if fc in ("03", "04"):
                name = "保持寄存器" if fc == "03" else "输入寄存器"
                self.append_log(
                    "TX",
                    f"读取{name}：unit={unit_id}, address={address}, count={count}",
                )
                if fc == "03":
                    regs = self._client.read_holding_registers(unit_id, address, count)
                else:
                    regs = self._client.read_input_registers(unit_id, address, count)
                self.append_log("RX", f"返回寄存器数量：{len(regs)}")
                self._fill_table(address, regs)
                self.append_log("OK", "操作完成")

            elif fc == "06":
                value = self._parse_write_values_single(raw_values)
                self.append_log(
                    "TX",
                    f"写单个保持寄存器：unit={unit_id}, address={address}, "
                    f"value={value} ({_fmt_reg_hex(value)})",
                )
                self._client.write_single_register(unit_id, address, value)
                self.append_log("RX", "写单个寄存器响应正常")
                self._clear_result_table()
                self.append_log("OK", "写入成功，结果表已清空。")

            elif fc == "16":
                values = self._parse_write_values_multi(raw_values)
                if count != len(values):
                    self._show_error(
                        "参数错误",
                        f"功能码 16 要求「数量」与写入值个数一致："
                        f"当前数量={count}，写入值个数={len(values)}。",
                    )
                    return
                self.append_log(
                    "TX",
                    f"写多个保持寄存器：unit={unit_id}, address={address}, "
                    f"count={len(values)}",
                )
                self._client.write_multiple_registers(unit_id, address, values)
                self.append_log("RX", f"写多个寄存器响应正常，共 {len(values)} 个寄存器")
                self._clear_result_table()
                self.append_log("OK", "写入成功，结果表已清空。")
            else:
                self._show_error("内部错误", f"未知功能码: {fc}")

        except ValueError as exc:
            self._show_error("参数错误", str(exc))
        except RuntimeError as exc:
            self._show_error("执行失败", self._format_user_exception(exc))
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            self.append_log("ERROR", f"未捕获异常:\n{tb}")
            self._show_error("执行失败", self._format_user_exception(exc))
        finally:
            self._set_exec_busy(False)

    # ------------------------------------------------------------------ 关闭窗口
    def closeEvent(self, event) -> None:  # type: ignore[override]
        self._client.close()
        super().closeEvent(event)
