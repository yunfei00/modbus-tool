"""
主窗口：连接配置、读写操作、结果表格、轮询、解析与日志。
所有 pymodbus 调用均通过 ModbusClient；解析与串口枚举在 core 子模块。
"""

from __future__ import annotations

import csv
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modbus_tool.core.config_manager import default_config_path, load_config, save_config
from modbus_tool.core.data_parser import build_parsed_column
from modbus_tool.core.modbus_client import ModbusClient
from modbus_tool.core.serial_utils import list_ports_available, list_serial_devices
from modbus_tool.version import APP_NAME, APP_VERSION


def _log_time() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _fmt_reg_hex(val: int) -> str:
    return f"0x{(val & 0xFFFF):04X}"


def _fmt_reg_bin(val: int) -> str:
    return "0b" + format(val & 0xFFFF, "016b")


@dataclass
class RegisterRowState:
    address: int
    current_value: int
    write_value: int
    status: str = "未修改"


class MainWindow(QWidget):
    """Modbus Studio 主界面。"""

    FC_OPTIONS = (
        ("03", "03 Read Holding Registers"),
        ("04", "04 Read Input Registers"),
        ("05", "05 Write Single Coil"),
        ("06", "06 Write Single Register"),
        ("0F", "0F Write Multiple Coils"),
        ("16", "16 Write Multiple Registers"),
    )

    LEFT_PANEL_WIDTH = 380
    OP_FIELD_MIN_WIDTH = 220
    MIN_POLL_INTERVAL_MS = 300

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} V3")
        self.resize(1280, 760)
        self.setMinimumSize(1000, 650)

        self._client = ModbusClient()
        self._exec_busy = False
        self._poll_busy = False
        self._polling_active = False
        self._last_registers: Optional[List[int]] = None
        self._last_start_addr: int = 0
        self._last_values_by_addr: Dict[int, int] = {}
        self._table_rows: List[RegisterRowState] = []
        self._updating_table: bool = False
        self._batch_mode: bool = False
        self._batch_addresses: List[int] = []
        self._manual_disconnect: bool = False
        self._last_conn_params: dict = {}
        self._reconnect_in_progress: bool = False
        self._throttle_warn_ts: float = 0.0
        self._success_count: int = 0
        self._failure_count: int = 0
        self._timeout_count: int = 0

        self._poll_timer = QTimer(self)
        self._poll_timer.timeout.connect(self._on_poll_tick)
        self._reconnect_timer = QTimer(self)
        self._reconnect_timer.setSingleShot(True)
        self._reconnect_timer.timeout.connect(self._attempt_reconnect)

        self._build_ui()
        self._wire_signals()

        # 默认 RTU：不依赖信号顺序，直接切换参数页
        self.combo_conn.blockSignals(True)
        self.combo_conn.setCurrentIndex(1)
        self.combo_conn.blockSignals(False)
        self._tcp_params_page.setVisible(False)
        self._rtu_params_page.setVisible(True)

        self._update_fc_dependent_widgets()
        self._update_poll_controls_enabled()
        self._set_connection_status(False)

        self.append_log(
            "INFO",
            f"{APP_NAME} V3 启动完成，默认通讯方式：RTU",
        )
        self.append_log("INFO", "结果表暂无数据，请执行读取操作。")
        self._refresh_serial_ports(log_info=False, is_startup=True)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        root = QHBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(10, 10, 10, 10)

        left_panel = QWidget()
        left_panel.setFixedWidth(self.LEFT_PANEL_WIDTH)
        left_panel.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Minimum)
        left_col = QVBoxLayout(left_panel)
        left_col.setSpacing(10)
        left_col.setContentsMargins(0, 0, 0, 0)

        # --- 配置（保存 / 加载）
        cfg_group = QGroupBox("配置")
        cfg_layout = QHBoxLayout(cfg_group)
        self.btn_save_cfg = QPushButton("保存配置")
        self.btn_load_cfg = QPushButton("加载配置")
        cfg_layout.addWidget(self.btn_save_cfg)
        cfg_layout.addWidget(self.btn_load_cfg)
        left_col.addWidget(cfg_group)

        # --- 连接配置
        conn_group = QGroupBox("连接配置")
        conn_layout = QVBoxLayout(conn_group)
        conn_layout.setSpacing(6)
        conn_layout.setContentsMargins(8, 8, 8, 8)

        row_type = QHBoxLayout()
        row_type.setSpacing(6)
        row_type.addWidget(QLabel("通讯类型:"))
        self.combo_conn = QComboBox()
        self.combo_conn.addItems(["TCP", "RTU"])
        row_type.addWidget(self.combo_conn)
        conn_layout.addLayout(row_type)

        self._params_holder = QWidget()
        ph_layout = QVBoxLayout(self._params_holder)
        ph_layout.setContentsMargins(0, 0, 0, 0)
        ph_layout.setSpacing(0)

        tcp_page = QWidget()
        tcp_form = QFormLayout(tcp_page)
        tcp_form.setSpacing(4)
        tcp_form.setContentsMargins(0, 2, 0, 2)
        tcp_form.setHorizontalSpacing(8)
        self.edit_host = QLineEdit("127.0.0.1")
        self.spin_port = QSpinBox()
        self.spin_port.setRange(1, 65535)
        self.spin_port.setValue(502)
        tcp_form.addRow("IP:", self.edit_host)
        tcp_form.addRow("端口:", self.spin_port)
        ph_layout.addWidget(tcp_page)

        rtu_page = QWidget()
        rtu_grid = QGridLayout(rtu_page)
        rtu_grid.setHorizontalSpacing(8)
        rtu_grid.setVerticalSpacing(4)
        rtu_grid.setContentsMargins(0, 2, 0, 2)
        self.combo_serial = QComboBox()
        self.combo_serial.setEditable(True)
        self.combo_serial.setMinimumContentsLength(10)
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
        self.combo_parity.setCurrentText("N")
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
        rtu_grid.setColumnStretch(1, 1)

        ph_layout.addWidget(rtu_page)
        self._tcp_params_page = tcp_page
        self._rtu_params_page = rtu_page
        conn_layout.addWidget(self._params_holder)

        row_unit = QHBoxLayout()
        row_unit.setSpacing(6)
        row_unit.addWidget(QLabel("从站地址:"))
        self.spin_unit = QSpinBox()
        self.spin_unit.setRange(1, 247)
        self.spin_unit.setValue(1)
        self.spin_unit.setMinimumWidth(72)
        row_unit.addWidget(self.spin_unit)
        row_unit.addStretch()
        conn_layout.addLayout(row_unit)

        row_btns = QHBoxLayout()
        row_btns.setSpacing(8)
        self.btn_connect = QPushButton("连接")
        self.btn_disconnect = QPushButton("断开")
        self.btn_disconnect.setEnabled(False)
        row_btns.addWidget(self.btn_connect)
        row_btns.addWidget(self.btn_disconnect)
        conn_layout.addLayout(row_btns)

        left_col.addWidget(conn_group)

        # --- 操作
        op_group = QGroupBox("操作")
        op_layout = QFormLayout(op_group)
        op_layout.setSpacing(4)
        op_layout.setContentsMargins(8, 8, 8, 8)
        op_layout.setHorizontalSpacing(8)

        self.combo_fc = QComboBox()
        for code, label in self.FC_OPTIONS:
            self.combo_fc.addItem(label, userData=code)
        self.combo_fc.setMinimumWidth(self.OP_FIELD_MIN_WIDTH)
        op_layout.addRow("功能码:", self.combo_fc)

        self.spin_addr = QSpinBox()
        self.spin_addr.setRange(0, 65535)
        self.spin_addr.setValue(0)
        self.spin_addr.setMinimumWidth(self.OP_FIELD_MIN_WIDTH)
        op_layout.addRow("起始地址:", self.spin_addr)

        self.edit_batch_addrs = QLineEdit()
        self.edit_batch_addrs.setMinimumWidth(self.OP_FIELD_MIN_WIDTH)
        self.edit_batch_addrs.setPlaceholderText("可选：批量地址，例如 0,1,5,10")
        op_layout.addRow("批量地址:", self.edit_batch_addrs)

        self.spin_count = QSpinBox()
        self.spin_count.setRange(1, 125)
        self.spin_count.setValue(10)
        self.spin_count.setMinimumWidth(self.OP_FIELD_MIN_WIDTH)
        op_layout.addRow("数量:", self.spin_count)

        self.edit_values = QLineEdit()
        self.edit_values.setMinimumWidth(self.OP_FIELD_MIN_WIDTH)
        op_layout.addRow("写入值:", self.edit_values)

        self.btn_exec = QPushButton("执行")
        self.btn_exec.setFixedHeight(34)
        self.btn_exec.setMinimumWidth(self.OP_FIELD_MIN_WIDTH)
        op_layout.addRow(self.btn_exec)

        action_row = QHBoxLayout()
        self.btn_read_once = QPushButton("一键读取")
        self.btn_clear_result = QPushButton("清空结果")
        self.btn_reset_config = QPushButton("重置配置")
        action_row.addWidget(self.btn_read_once)
        action_row.addWidget(self.btn_clear_result)
        action_row.addWidget(self.btn_reset_config)
        op_layout.addRow(action_row)

        left_col.addWidget(op_group)

        # --- 轮询
        poll_group = QGroupBox("周期轮询")
        poll_layout = QVBoxLayout(poll_group)
        self.chk_poll = QCheckBox("启用轮询")
        poll_layout.addWidget(self.chk_poll)
        row_pi = QHBoxLayout()
        row_pi.addWidget(QLabel("间隔 (ms):"))
        self.spin_poll_interval = QSpinBox()
        self.spin_poll_interval.setRange(self.MIN_POLL_INTERVAL_MS, 600_000)
        self.spin_poll_interval.setValue(1000)
        self.spin_poll_interval.setMinimumWidth(120)
        row_pi.addWidget(self.spin_poll_interval)
        row_pi.addStretch()
        poll_layout.addLayout(row_pi)
        row_pb = QHBoxLayout()
        self.btn_poll_start = QPushButton("开始轮询")
        self.btn_poll_stop = QPushButton("停止轮询")
        self.btn_poll_stop.setEnabled(False)
        row_pb.addWidget(self.btn_poll_start)
        row_pb.addWidget(self.btn_poll_stop)
        poll_layout.addLayout(row_pb)
        left_col.addWidget(poll_group)

        stats_group = QGroupBox("错误统计")
        stats_layout = QFormLayout(stats_group)
        stats_layout.setContentsMargins(8, 8, 8, 8)
        self.lbl_success_count = QLabel("0")
        self.lbl_failure_count = QLabel("0")
        self.lbl_timeout_count = QLabel("0")
        stats_layout.addRow("成功次数:", self.lbl_success_count)
        stats_layout.addRow("失败次数:", self.lbl_failure_count)
        stats_layout.addRow("超时次数:", self.lbl_timeout_count)
        left_col.addWidget(stats_group)

        # --- 连接状态
        status_group = QGroupBox("连接状态")
        status_layout = QHBoxLayout(status_group)
        status_layout.setContentsMargins(8, 6, 8, 6)
        self.lbl_status_icon = QLabel("●")
        self.lbl_status_icon.setFixedWidth(22)
        self.lbl_status_icon.setAlignment(Qt.AlignCenter)
        self.lbl_status_text = QLabel("状态：未连接")
        self.lbl_status_text.setMinimumHeight(36)
        self.lbl_status_text.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.lbl_status_icon)
        status_layout.addWidget(self.lbl_status_text, stretch=1)
        left_col.addWidget(status_group)

        root.addWidget(left_panel, stretch=0, alignment=Qt.AlignmentFlag.AlignTop)

        # --- 右侧
        right_panel = QWidget()
        right_col = QVBoxLayout(right_panel)
        right_col.setSpacing(8)

        splitter = QSplitter(Qt.Vertical)

        res_group = QGroupBox("结果")
        res_layout = QVBoxLayout(res_group)

        parse_row = QHBoxLayout()
        parse_row.addWidget(QLabel("数据类型:"))
        self.combo_dtype = QComboBox()
        self.combo_dtype.addItems(["uint16", "int16", "uint32", "int32", "float32"])
        parse_row.addWidget(self.combo_dtype)
        parse_row.addWidget(QLabel("字序:"))
        self.combo_endian = QComboBox()
        self.combo_endian.addItems(["AB CD", "CD AB"])
        parse_row.addWidget(self.combo_endian)
        parse_row.addStretch()
        self.btn_write_selected = QPushButton("写回选中行")
        self.btn_write_all_modified = QPushButton("写回所有修改")
        self.btn_reread = QPushButton("重新读取")
        self.btn_clear_result_table = QPushButton("清空结果")
        self.btn_write_all_modified.setStyleSheet("font-weight: 700;")
        parse_row.addWidget(self.btn_write_selected)
        parse_row.addWidget(self.btn_write_all_modified)
        parse_row.addWidget(self.btn_reread)
        parse_row.addWidget(self.btn_clear_result_table)
        self.btn_export_csv = QPushButton("导出 CSV")
        parse_row.addWidget(self.btn_export_csv)
        res_layout.addLayout(parse_row)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(
            ["地址", "当前值(DEC)", "当前值(HEX)", "当前值(BIN)", "解析值", "写入值", "状态"]
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.SelectedClicked
            | QAbstractItemView.EditKeyPressed
        )
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.setStyleSheet("QTableWidget { font-size: 13px; }")
        res_layout.addWidget(self.table)
        splitter.addWidget(res_group)

        log_group = QGroupBox("通信日志")
        log_outer = QVBoxLayout(log_group)
        log_btns = QHBoxLayout()
        self.btn_log_clear = QPushButton("清空日志")
        self.btn_log_save = QPushButton("保存日志")
        log_btns.addWidget(self.btn_log_clear)
        log_btns.addWidget(self.btn_log_save)
        log_btns.addStretch()
        log_outer.addLayout(log_btns)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(200)
        _mono = QFont("Consolas", 10)
        if not _mono.exactMatch():
            _mono = QFont("Courier New", 10)
        self.log.setFont(_mono)
        self._apply_button_style()
        log_outer.addWidget(self.log)
        splitter.addWidget(log_group)

        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 3)
        splitter.setSizes([520, 280])
        right_col.addWidget(splitter)
        root.addWidget(right_panel, stretch=1)

    def _apply_button_style(self) -> None:
        for btn in self.findChildren(QPushButton):
            btn.setMinimumHeight(34)

    def _wire_signals(self) -> None:
        self.combo_conn.currentIndexChanged.connect(self._on_conn_type_changed)
        self.btn_refresh_ports.clicked.connect(self._refresh_serial_ports)
        self.btn_connect.clicked.connect(self._on_connect)
        self.btn_disconnect.clicked.connect(self._on_disconnect)
        self.btn_exec.clicked.connect(self._on_execute)
        self.btn_read_once.clicked.connect(self._on_read_once)
        self.btn_clear_result.clicked.connect(self._on_clear_result)
        self.btn_reset_config.clicked.connect(self._on_reset_config)
        self.combo_fc.currentIndexChanged.connect(self._on_fc_changed)
        self.combo_dtype.currentTextChanged.connect(self._on_parse_options_changed)
        self.combo_endian.currentTextChanged.connect(self._on_parse_options_changed)
        self.btn_save_cfg.clicked.connect(self._on_save_config)
        self.btn_load_cfg.clicked.connect(self._on_load_config)
        self.btn_log_clear.clicked.connect(self._on_clear_log)
        self.btn_log_save.clicked.connect(self._on_save_log)
        self.btn_export_csv.clicked.connect(self._on_export_csv)
        self.btn_write_selected.clicked.connect(self._on_write_selected_row)
        self.btn_write_all_modified.clicked.connect(self._on_write_all_modified)
        self.btn_reread.clicked.connect(self._on_reread)
        self.btn_clear_result_table.clicked.connect(self._on_clear_result)
        self.table.itemChanged.connect(self._on_table_item_changed)
        self.table.itemSelectionChanged.connect(self._update_table_write_controls)
        self.btn_poll_start.clicked.connect(self._on_poll_start)
        self.btn_poll_stop.clicked.connect(self._on_poll_stop)
        self.chk_poll.stateChanged.connect(self._update_poll_controls_enabled)
        self.spin_poll_interval.valueChanged.connect(
            lambda *_: self._update_poll_controls_enabled()
        )

    # ------------------------------------------------------------------ 日志
    def append_log(self, level: str, message: str) -> None:
        level = level.upper()
        self.log.append(f"[{_log_time()}] [{level}] {message}")
        sb = self.log.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _show_error(self, title: str, message: str) -> None:
        self.append_log("ERROR", f"{title}: {message}")
        QMessageBox.warning(self, title, message)

    def _show_warning(self, title: str, message: str) -> None:
        self.append_log("WARN", f"{title}: {message}")
        QMessageBox.warning(self, title, message)

    def _warn_not_connected_execute(self) -> None:
        self.append_log("WARN", "未连接：请先连接设备。")
        QMessageBox.warning(self, "未连接", "请先连接设备")

    def _set_exec_busy(self, busy: bool) -> None:
        self._exec_busy = busy
        self.btn_exec.setEnabled(not busy)
        self.btn_exec.setText("执行中..." if busy else "执行")

    def _set_connection_status(self, connected: bool) -> None:
        if connected:
            self.lbl_status_icon.setStyleSheet("color: #43A047; font-size: 22px;")
            self.lbl_status_text.setText("状态：已连接")
            self.lbl_status_text.setStyleSheet(
                "color: #1B5E20; font-size: 15px; font-weight: 700; "
                "background: #E8F5E9; border: 1px solid #81C784; border-radius: 6px;"
            )
        else:
            self.lbl_status_icon.setStyleSheet("color: #E53935; font-size: 22px;")
            self.lbl_status_text.setText("状态：未连接")
            self.lbl_status_text.setStyleSheet(
                "color: #B71C1C; font-size: 15px; font-weight: 700; "
                "background: #FFEBEE; border: 1px solid #EF9A9A; border-radius: 6px;"
            )

    def _refresh_stats_labels(self) -> None:
        self.lbl_success_count.setText(str(self._success_count))
        self.lbl_failure_count.setText(str(self._failure_count))
        self.lbl_timeout_count.setText(str(self._timeout_count))

    def _record_success(self) -> None:
        self._success_count += 1
        self._refresh_stats_labels()
        self.append_log(
            "INFO",
            f"统计更新：成功={self._success_count} 失败={self._failure_count} 超时={self._timeout_count}",
        )

    def _record_failure(self, exc: BaseException) -> None:
        self._failure_count += 1
        txt = str(exc).lower()
        if "timeout" in txt or "timed out" in txt or isinstance(exc, TimeoutError):
            self._timeout_count += 1
        self._refresh_stats_labels()
        self.append_log(
            "WARN",
            f"统计更新：成功={self._success_count} 失败={self._failure_count} 超时={self._timeout_count}",
        )

    # ------------------------------------------------------------------ 串口
    def _refresh_serial_ports(self, log_info: bool = True, is_startup: bool = False) -> None:
        current = self.combo_serial.currentText().strip()
        self.combo_serial.clear()

        if not list_ports_available():
            self.append_log(
                "WARN",
                "无法使用 serial.tools.list_ports，请检查 pyserial 安装或手动输入串口名。",
            )
            if current:
                self.combo_serial.addItem(current)
                self.combo_serial.setEditText(current)
            return

        devices = list_serial_devices()
        for dev in devices:
            self.combo_serial.addItem(dev, userData=dev)

        if not devices:
            self.append_log("WARN", "未检测到可用串口")
            if current:
                self.combo_serial.addItem(current)
                self.combo_serial.setEditText(current)
            return

        idx = self.combo_serial.findText(current)
        if idx >= 0:
            self.combo_serial.setCurrentIndex(idx)
        elif current:
            self.combo_serial.setEditText(current)
        else:
            self.combo_serial.setCurrentIndex(0)

        if log_info and not is_startup:
            self.append_log("INFO", f"串口列表已刷新，共 {len(devices)} 个端口。")

    def _current_serial_device(self) -> str:
        i = self.combo_serial.currentIndex()
        if i >= 0:
            data = self.combo_serial.itemData(i)
            if data:
                return str(data)
        return self.combo_serial.currentText().strip()

    def _on_conn_type_changed(self, _index: int) -> None:
        is_tcp = self.combo_conn.currentText() == "TCP"
        self._tcp_params_page.setVisible(is_tcp)
        self._rtu_params_page.setVisible(not is_tcp)

    def _current_fc(self) -> str:
        return str(self.combo_fc.currentData())

    def _on_fc_changed(self, _index: int) -> None:
        fc = self._current_fc()
        if fc in ("06", "16") and self._polling_active:
            self.append_log("WARN", "写操作不支持轮询，已自动停止轮询。")
            self._stop_polling_internal()
        self._update_fc_dependent_widgets()
        self._update_poll_controls_enabled()

    def _update_fc_dependent_widgets(self) -> None:
        fc = self._current_fc()
        is_read = fc in ("03", "04")
        is_single = fc in ("05", "06")
        is_multi = fc in ("0F", "16")

        if is_single:
            self.spin_count.setValue(1)
            self.spin_count.setEnabled(False)
        else:
            self.spin_count.setEnabled(is_read or is_multi)

        self.edit_values.setEnabled(is_single or is_multi)
        self.edit_batch_addrs.setEnabled(is_read)

        if is_read:
            self.edit_values.clear()
            self.edit_values.setPlaceholderText("读操作无需填写写入值")
            self.edit_batch_addrs.setPlaceholderText("可选：批量地址，例如 0,1,5,10")
        elif fc == "05":
            self.edit_values.setPlaceholderText("请输入线圈值，例如 1/0、true/false、on/off")
            self.edit_batch_addrs.clear()
            self.edit_batch_addrs.setPlaceholderText("仅读操作支持批量地址")
        elif fc == "06":
            self.edit_values.setPlaceholderText("请输入单个寄存器值，例如 123")
            self.edit_batch_addrs.clear()
            self.edit_batch_addrs.setPlaceholderText("仅读操作支持批量地址")
        elif fc == "0F":
            self.edit_values.setPlaceholderText("请输入多个线圈值，例如 1,0,1 或 true,false,true")
            self.edit_batch_addrs.clear()
            self.edit_batch_addrs.setPlaceholderText("仅读操作支持批量地址")
        else:
            self.edit_values.setPlaceholderText("请输入多个寄存器值，例如 1,2,3,4")
            self.edit_batch_addrs.clear()
            self.edit_batch_addrs.setPlaceholderText("仅读操作支持批量地址")
        self._updating_table = True
        for row_idx in range(len(self._table_rows)):
            self._refresh_table_row(row_idx)
        self._updating_table = False
        self._update_table_write_controls()

    def _update_poll_controls_enabled(self) -> None:
        fc_ok = self._current_fc() in ("03", "04")
        self.chk_poll.setEnabled(fc_ok)
        self.spin_poll_interval.setEnabled(fc_ok and not self._polling_active)
        can_start = (
            fc_ok
            and self.chk_poll.isChecked()
            and not self._polling_active
            and self._client.is_connected()
        )
        self.btn_poll_start.setEnabled(can_start)
        self.btn_poll_stop.setEnabled(self._polling_active and fc_ok)

    def _update_table_write_controls(self) -> None:
        is_holding_read = self._current_fc() == "03"
        has_rows = self.table.rowCount() > 0
        has_selection = bool(self.table.selectionModel() and self.table.selectionModel().selectedRows())
        pending_count = sum(1 for row in self._table_rows if row.status == "待写入")
        self.btn_write_selected.setEnabled(is_holding_read and has_rows and has_selection)
        self.btn_write_all_modified.setEnabled(is_holding_read and pending_count > 0)
        self.btn_reread.setEnabled(has_rows and self._current_fc() in ("03", "04"))
        if not is_holding_read:
            self.btn_write_selected.setToolTip("仅功能码 03 结果支持表格写回")
            self.btn_write_all_modified.setToolTip("仅功能码 03 结果支持表格写回")
        else:
            self.btn_write_selected.setToolTip("")
            self.btn_write_all_modified.setToolTip("")

    def _set_connected_ui(self, connected: bool) -> None:
        self.btn_connect.setEnabled(not connected)
        self.btn_disconnect.setEnabled(connected)
        self.combo_conn.setEnabled(not connected)
        self._params_holder.setEnabled(not connected)
        self.spin_unit.setEnabled(not connected)
        self.btn_refresh_ports.setEnabled(not connected)
        self._set_connection_status(connected)
        self._update_poll_controls_enabled()
        self._update_table_write_controls()

    def _stop_polling_internal(self, silent: bool = False) -> None:
        was_active = self._polling_active
        self._poll_timer.stop()
        self._polling_active = False
        if was_active and not silent:
            self.append_log("INFO", "轮询已停止。")
        self._update_poll_controls_enabled()

    def _on_poll_start(self) -> None:
        if not self._client.is_connected():
            self._show_warning("未连接", "请先连接设备再开始轮询。")
            return
        if self._current_fc() not in ("03", "04"):
            self._show_warning("轮询", "仅支持功能码 03 / 04 轮询。")
            return
        if not self.chk_poll.isChecked():
            self._show_warning("轮询", "请先勾选「启用轮询」。")
            return
        if int(self.spin_poll_interval.value()) < self.MIN_POLL_INTERVAL_MS:
            self.spin_poll_interval.setValue(self.MIN_POLL_INTERVAL_MS)
            self.append_log("WARN", f"轮询间隔过小，已调整为 {self.MIN_POLL_INTERVAL_MS} ms")
        self._polling_active = True
        self._poll_timer.setInterval(int(self.spin_poll_interval.value()))
        self._poll_timer.start()
        self.btn_poll_start.setEnabled(False)
        self.btn_poll_stop.setEnabled(True)
        self.spin_poll_interval.setEnabled(False)
        self.append_log(
            "INFO",
            f"开始轮询：间隔 {self.spin_poll_interval.value()} ms，功能码 {self._current_fc()}",
        )

    def _on_poll_stop(self) -> None:
        self._stop_polling_internal()

    def _on_poll_tick(self) -> None:
        if not self._polling_active or not self._client.is_connected():
            if self._polling_active and not self._client.is_connected():
                self._schedule_reconnect("轮询检测到连接断开")
            return
        if self._poll_busy or self._exec_busy:
            now = time.monotonic()
            if now - self._throttle_warn_ts > 3:
                self.append_log("INFO", "请求节流：上一次请求未完成，跳过本次轮询。")
                self._throttle_warn_ts = now
            return
        if self._current_fc() not in ("03", "04"):
            self._stop_polling_internal()
            return
        self._poll_busy = True
        try:
            self._perform_read_and_fill(log_tx_rx=True, log_ok=False)
            self._record_success()
        except Exception as exc:  # noqa: BLE001
            self._record_failure(exc)
            self.append_log("ERROR", f"轮询失败: {self._format_user_exception(exc)}")
            if self._is_connection_like_error(exc):
                self._schedule_reconnect(str(exc))
        finally:
            self._poll_busy = False

    # ------------------------------------------------------------------ 连接
    def _on_connect(self) -> None:
        if self._client.is_connected():
            self._show_warning("提示", "已经处于连接状态。")
            return

        try:
            self._manual_disconnect = False
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
                    f"正在连接 RTU 串口 {port_name}，{baud}/{bytesize}{parity}{stop}",
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
        self._store_connection_params()
        self._reconnect_in_progress = False

    def _on_disconnect(self) -> None:
        self._manual_disconnect = True
        self._reconnect_timer.stop()
        self._reconnect_in_progress = False
        self._stop_polling_internal(silent=True)
        self._client.close()
        self._set_connected_ui(False)
        self.append_log("INFO", "已断开连接（轮询已停止）。")

    # ------------------------------------------------------------------ 解析 / 表格
    def _on_parse_options_changed(self, *_args: object) -> None:
        self._apply_parsed_column()

    def _apply_parsed_column(self) -> None:
        if self._last_registers is None or self.table.rowCount() == 0:
            return
        dtype = self.combo_dtype.currentText()
        order = self.combo_endian.currentText()
        try:
            parsed = build_parsed_column(self._last_registers, dtype, order)
        except Exception as exc:  # noqa: BLE001
            self.append_log("WARN", f"解析选项应用失败: {exc}")
            return
        for row in range(min(self.table.rowCount(), len(parsed))):
            self.table.setItem(row, 4, QTableWidgetItem(parsed[row]))

    def _clear_result_table(self) -> None:
        self._updating_table = True
        self.table.setRowCount(0)
        self._updating_table = False
        self._last_registers = None
        self._last_values_by_addr.clear()
        self._table_rows.clear()
        self.append_log("INFO", "结果表暂无数据，请执行读取操作。")
        self._update_table_write_controls()

    def _set_table_item(self, row: int, col: int, text: str, editable: bool = False) -> None:
        item = QTableWidgetItem(text)
        flags = Qt.ItemIsSelectable | Qt.ItemIsEnabled
        if editable:
            flags |= Qt.ItemIsEditable
            item.setBackground(QColor("#FFF8E1"))
        item.setFlags(flags)
        self.table.setItem(row, col, item)

    def _status_text(self, row: RegisterRowState, edited: bool = False) -> str:
        if row.status == "写入失败":
            return "写入失败"
        if row.status == "写入成功":
            return "写入成功"
        if edited or row.write_value != row.current_value:
            return "待写入"
        return "未修改"

    def _refresh_table_row(self, row_index: int) -> None:
        if row_index >= len(self._table_rows):
            return
        row = self._table_rows[row_index]
        self._set_table_item(row_index, 0, str(row.address))
        self._set_table_item(row_index, 1, str(row.current_value))
        self._set_table_item(row_index, 2, _fmt_reg_hex(row.current_value))
        self._set_table_item(row_index, 3, _fmt_reg_bin(row.current_value))
        self._set_table_item(row_index, 4, "")
        self._set_table_item(row_index, 5, str(row.write_value), editable=self._current_fc() == "03")
        self._set_table_item(row_index, 6, row.status)

    def _fill_table(self, start_addr: int, registers: List[int]) -> None:
        self._last_start_addr = start_addr
        self._last_registers = list(registers)
        self._table_rows = []
        self._updating_table = True
        self.table.setRowCount(0)
        for i, val in enumerate(registers):
            row = self.table.rowCount()
            self.table.insertRow(row)
            addr = start_addr + i
            row_state = RegisterRowState(address=addr, current_value=val, write_value=val)
            self._table_rows.append(row_state)
            self._refresh_table_row(row)
            self._last_values_by_addr[addr] = val
        self._updating_table = False
        self._apply_parsed_column()
        self._update_table_write_controls()

    def _fill_batch_table(self, values_by_addr: Dict[int, int]) -> None:
        self._updating_table = True
        self.table.setRowCount(0)
        self._last_registers = list(values_by_addr.values())
        self._table_rows = []
        for addr in sorted(values_by_addr.keys()):
            val = values_by_addr[addr]
            prev = self._last_values_by_addr.get(addr)
            row = self.table.rowCount()
            self.table.insertRow(row)
            row_state = RegisterRowState(address=addr, current_value=val, write_value=val)
            self._table_rows.append(row_state)
            self._refresh_table_row(row)
            parsed_item = self.table.item(row, 4)
            if prev is not None and val != prev:
                for item in (
                    self.table.item(row, 1),
                    self.table.item(row, 2),
                    self.table.item(row, 3),
                    parsed_item,
                ):
                    item.setForeground(Qt.GlobalColor.red if val > prev else Qt.GlobalColor.blue)
                self.append_log(
                    "INFO",
                    f"地址 {addr} 数值变化：{prev} -> {val}（{'变大' if val > prev else '变小'}）",
                )
            self._last_values_by_addr[addr] = val
        self._updating_table = False
        self._update_table_write_controls()

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._updating_table:
            return
        row = item.row()
        col = item.column()
        if col != 5 or row >= len(self._table_rows):
            return
        row_state = self._table_rows[row]
        old_value = row_state.write_value
        raw = (item.text() or "").strip()
        try:
            new_value = int(raw, 10)
        except ValueError:
            self._show_warning("输入错误", "写入值必须是整数，范围 0~65535。")
            self._updating_table = True
            item.setText(str(old_value))
            self._updating_table = False
            return
        if not 0 <= new_value <= 65535:
            self._show_warning("输入错误", "写入值必须在 0~65535 范围内。")
            self._updating_table = True
            item.setText(str(old_value))
            self._updating_table = False
            return
        row_state.write_value = new_value
        row_state.status = self._status_text(row_state, edited=True)
        self._updating_table = True
        self.table.item(row, 6).setText(row_state.status)
        self._updating_table = False
        self._update_table_write_controls()

    def _read_addresses_holding(self, addresses: List[int]) -> Dict[int, int]:
        if not addresses:
            return {}
        unit_id = int(self.spin_unit.value())
        result: Dict[int, int] = {}
        sorted_addrs = sorted(addresses)
        if sorted_addrs == list(range(sorted_addrs[0], sorted_addrs[-1] + 1)):
            start = sorted_addrs[0]
            count = len(sorted_addrs)
            regs = self._client.read_holding_registers(unit_id, start, count)
            for idx, value in enumerate(regs):
                result[start + idx] = value
            return result
        for addr in sorted_addrs:
            result[addr] = self._client.read_holding_registers(unit_id, addr, 1)[0]
        return result

    def _sync_rows_after_reread(self, values_by_addr: Dict[int, int]) -> None:
        self._updating_table = True
        for idx, row_state in enumerate(self._table_rows):
            if row_state.address not in values_by_addr:
                continue
            row_state.current_value = values_by_addr[row_state.address]
            if row_state.status != "待写入":
                row_state.write_value = row_state.current_value
                row_state.status = "未修改"
            self._refresh_table_row(idx)
            self._last_values_by_addr[row_state.address] = row_state.current_value
        self._updating_table = False
        self._last_registers = [row.current_value for row in self._table_rows]
        self._apply_parsed_column()
        self._update_table_write_controls()

    # ------------------------------------------------------------------ 读 / 执行
    def _parse_batch_addresses(self) -> List[int]:
        text = self.edit_batch_addrs.text().strip()
        if not text:
            return []
        values: List[int] = []
        seen: set[int] = set()
        for part in text.split(","):
            s = part.strip()
            if not s:
                continue
            addr = int(s, 0)
            if not 0 <= addr <= 65535:
                raise ValueError(f"地址超出范围(0~65535): {s}")
            if addr not in seen:
                seen.add(addr)
                values.append(addr)
        if not values:
            raise ValueError("批量地址不能为空，请使用英文逗号分隔，例如 0,1,5")
        return values

    def _do_read_single_address(self, fc: str, unit_id: int, address: int) -> int:
        if fc == "03":
            return self._client.read_holding_registers(unit_id, address, 1)[0]
        return self._client.read_input_registers(unit_id, address, 1)[0]

    def _is_connection_like_error(self, exc: BaseException) -> bool:
        txt = str(exc).lower()
        marks = ("timeout", "timed out", "connection", "no response", "socket", "通讯异常")
        return any(mark in txt for mark in marks)

    def _store_connection_params(self) -> None:
        if self.combo_conn.currentText() == "TCP":
            self._last_conn_params = {
                "kind": "TCP",
                "host": self.edit_host.text().strip(),
                "port": int(self.spin_port.value()),
            }
        else:
            self._last_conn_params = {
                "kind": "RTU",
                "port": self._current_serial_device(),
                "baudrate": int(self.spin_baud.value()),
                "bytesize": int(self.spin_bytesize.currentText()),
                "parity": self.combo_parity.currentText(),
                "stopbits": float(self.combo_stopbits.currentText()),
            }

    def _schedule_reconnect(self, reason: str) -> None:
        if self._manual_disconnect:
            return
        if self._reconnect_in_progress:
            return
        self._reconnect_in_progress = True
        self._set_connected_ui(False)
        self.append_log("WARN", f"连接中断，准备自动重连：{reason}")
        self._reconnect_timer.start(1500)

    def _attempt_reconnect(self) -> None:
        if not self._last_conn_params:
            self._reconnect_in_progress = False
            self.append_log("WARN", "缺少连接参数，无法自动重连。")
            return
        try:
            self.append_log("INFO", "自动重连中...")
            if self._last_conn_params.get("kind") == "TCP":
                self._client.connect_tcp(
                    str(self._last_conn_params["host"]),
                    int(self._last_conn_params["port"]),
                )
            else:
                self._client.connect_rtu(
                    str(self._last_conn_params["port"]),
                    int(self._last_conn_params["baudrate"]),
                    int(self._last_conn_params["bytesize"]),
                    str(self._last_conn_params["parity"]),
                    float(self._last_conn_params["stopbits"]),
                )
            self._set_connected_ui(True)
            self.append_log("OK", "自动重连成功。")
            self._reconnect_in_progress = False
        except Exception as exc:  # noqa: BLE001
            self.append_log("WARN", f"自动重连失败，将继续重试：{self._format_user_exception(exc)}")
            self._reconnect_timer.start(2000)

    def _perform_read_and_fill(self, log_tx_rx: bool, log_ok: bool) -> None:
        unit_id = int(self.spin_unit.value())
        fc = self._current_fc()
        if fc not in ("03", "04"):
            raise RuntimeError("仅支持读保持/输入寄存器")
        name = "保持寄存器" if fc == "03" else "输入寄存器"
        batch_addrs = self._parse_batch_addresses()
        self._batch_mode = bool(batch_addrs)
        self._batch_addresses = list(batch_addrs)
        if batch_addrs:
            if log_tx_rx:
                self.append_log(
                    "TX",
                    f"批量读取{name}：unit={unit_id}, addresses={batch_addrs}",
                )
            values_by_addr: Dict[int, int] = {}
            for addr in batch_addrs:
                try:
                    values_by_addr[addr] = self._do_read_single_address(fc, unit_id, addr)
                    self.append_log("RX", f"地址 {addr} 读取成功，值={values_by_addr[addr]}")
                except Exception as exc:  # noqa: BLE001
                    self.append_log("ERROR", f"地址 {addr} 读取失败：{self._format_user_exception(exc)}")
                    if self._is_connection_like_error(exc):
                        self._schedule_reconnect(str(exc))
            self._fill_batch_table(values_by_addr)
        else:
            address = int(self.spin_addr.value())
            count = int(self.spin_count.value())
            if log_tx_rx:
                self.append_log(
                    "TX",
                    f"读取{name}：unit={unit_id}, address={address}, count={count}",
                )
            if fc == "03":
                regs = self._client.read_holding_registers(unit_id, address, count)
            else:
                regs = self._client.read_input_registers(unit_id, address, count)
            if log_tx_rx:
                self.append_log("RX", f"返回寄存器数量：{len(regs)}")
            self._fill_table(address, regs)
        if log_ok:
            self.append_log("OK", "操作完成")

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

    def _parse_coil_value(self, text: str) -> bool:
        s = text.strip().lower()
        if not s:
            raise ValueError("线圈写入值不能为空。")
        truthy = {"1", "true", "on", "yes"}
        falsy = {"0", "false", "off", "no"}
        if s in truthy:
            return True
        if s in falsy:
            return False
        raise ValueError("线圈值仅支持 1/0、true/false、on/off、yes/no。")

    def _parse_coil_values_multi(self, text: str) -> list[bool]:
        s = text.strip()
        if not s:
            raise ValueError("线圈写入值不能为空。")
        parts = [p.strip() for p in s.split(",") if p.strip()]
        if not parts:
            raise ValueError("请使用英文逗号分隔多个线圈值，例如: 1,0,1")
        return [self._parse_coil_value(p) for p in parts]

    def _format_user_exception(self, exc: BaseException, context: str = "") -> str:
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

    def _on_execute(self) -> None:
        if not self._client.is_connected():
            self._warn_not_connected_execute()
            return

        self._set_exec_busy(True)
        try:
            unit_id = int(self.spin_unit.value())
            address = int(self.spin_addr.value())
            count = int(self.spin_count.value())
            fc = self._current_fc()
            raw_values = self.edit_values.text()

            if fc in ("03", "04"):
                self._perform_read_and_fill(log_tx_rx=True, log_ok=True)
                self._record_success()

            elif fc == "05":
                value = self._parse_coil_value(raw_values)
                self.append_log(
                    "TX",
                    f"写单个线圈：unit={unit_id}, address={address}, value={int(value)}",
                )
                self._client.write_single_coil(unit_id, address, value)
                self.append_log("RX", "写单个线圈响应正常")
                self._clear_result_table()
                self.append_log("OK", "写入成功。")
                self._record_success()

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
                self.append_log("OK", "写入成功。")
                self._record_success()

            elif fc == "0F":
                values = self._parse_coil_values_multi(raw_values)
                if count != len(values):
                    self._show_error(
                        "参数错误",
                        f"功能码 0F 要求「数量」与写入值个数一致："
                        f"当前数量={count}，写入值个数={len(values)}。",
                    )
                    return
                self.append_log(
                    "TX",
                    f"写多个线圈：unit={unit_id}, address={address}, count={len(values)}",
                )
                self._client.write_multiple_coils(unit_id, address, values)
                self.append_log("RX", f"写多个线圈响应正常，共 {len(values)} 个线圈")
                self._clear_result_table()
                self.append_log("OK", "写入成功。")
                self._record_success()

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
                self.append_log("OK", "写入成功。")
                self._record_success()
            else:
                self._show_error("内部错误", f"未知功能码: {fc}")

        except ValueError as exc:
            self._record_failure(exc)
            self._show_error("参数错误", str(exc))
        except RuntimeError as exc:
            self._record_failure(exc)
            self._show_error("执行失败", self._format_user_exception(exc))
            if self._is_connection_like_error(exc):
                self._schedule_reconnect(str(exc))
        except Exception as exc:  # noqa: BLE001
            self._record_failure(exc)
            tb = traceback.format_exc()
            self.append_log("ERROR", f"未捕获异常:\n{tb}")
            self._show_error("执行失败", self._format_user_exception(exc))
            if self._is_connection_like_error(exc):
                self._schedule_reconnect(str(exc))
        finally:
            self._set_exec_busy(False)

    def _on_read_once(self) -> None:
        if self._current_fc() not in ("03", "04"):
            self.append_log("WARN", "一键读取仅支持功能码 03 / 04。")
            return
        self.append_log("INFO", "触发一键读取。")
        self._on_execute()

    def _on_reread(self) -> None:
        if not self._client.is_connected():
            self._warn_not_connected_execute()
            return
        if self._current_fc() not in ("03", "04"):
            self._show_warning("重新读取", "当前功能码不支持重新读取。")
            return
        if not self._table_rows:
            self._show_warning("重新读取", "结果表暂无数据。")
            return
        try:
            unit_id = int(self.spin_unit.value())
            fc = self._current_fc()
            values_by_addr: Dict[int, int] = {}
            batch_addrs = self._parse_batch_addresses()
            if batch_addrs:
                self.append_log("TX", f"重新读取批量地址：{batch_addrs}")
                for addr in batch_addrs:
                    if fc == "03":
                        values_by_addr[addr] = self._client.read_holding_registers(unit_id, addr, 1)[0]
                    else:
                        values_by_addr[addr] = self._client.read_input_registers(unit_id, addr, 1)[0]
            else:
                start = int(self.spin_addr.value())
                count = int(self.spin_count.value())
                self.append_log("TX", f"重新读取范围：address={start}, count={count}")
                if fc == "03":
                    regs = self._client.read_holding_registers(unit_id, start, count)
                else:
                    regs = self._client.read_input_registers(unit_id, start, count)
                for idx, reg in enumerate(regs):
                    values_by_addr[start + idx] = reg
            self._sync_rows_after_reread(values_by_addr)
            self.append_log("OK", "重新读取完成。")
        except Exception as exc:  # noqa: BLE001
            self._show_error("重新读取失败", self._format_user_exception(exc))

    def _on_write_selected_row(self) -> None:
        if self._current_fc() != "03":
            self._show_warning("写回", "仅功能码 03 结果支持写回。")
            return
        if not self._client.is_connected():
            self._warn_not_connected_execute()
            return
        selected = self.table.selectionModel().selectedRows() if self.table.selectionModel() else []
        if not selected:
            self._show_warning("写回选中行", "请先选中一行。")
            return
        row_idx = selected[0].row()
        row_state = self._table_rows[row_idx]
        ok = QMessageBox.question(
            self,
            "确认写入",
            "确认写入当前选中寄存器？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ok != QMessageBox.Yes:
            return
        unit_id = int(self.spin_unit.value())
        expected = row_state.write_value
        addr = row_state.address
        try:
            self.append_log("TX", f"写单寄存器：address={addr}, value={expected}")
            self._client.write_single_register(unit_id, addr, expected)
            self.append_log("OK", f"写入成功：address={addr}")
            actual = self._client.read_holding_registers(unit_id, addr, 1)[0]
            self.append_log("RX", f"重读校验：address={addr}, value={actual}")
            row_state.current_value = actual
            if actual == expected:
                row_state.status = "写入成功"
            else:
                row_state.status = "写入失败"
                self.append_log(
                    "ERROR",
                    f"写入失败：address={addr}, expected={expected}, actual={actual}",
                )
            row_state.write_value = expected
            self._updating_table = True
            self._refresh_table_row(row_idx)
            self._updating_table = False
            self._apply_parsed_column()
            self._update_table_write_controls()
        except Exception as exc:  # noqa: BLE001
            row_state.status = "写入失败"
            self._updating_table = True
            self.table.item(row_idx, 6).setText("写入失败")
            self._updating_table = False
            self.append_log("ERROR", f"写入失败：address={addr}, reason={self._format_user_exception(exc)}")
            self._show_error("写回失败", self._format_user_exception(exc))

    def _on_write_all_modified(self) -> None:
        if self._current_fc() != "03":
            self._show_warning("批量写回", "仅功能码 03 结果支持写回。")
            return
        if not self._client.is_connected():
            self._warn_not_connected_execute()
            return
        pending_rows = [row for row in self._table_rows if row.status == "待写入"]
        if not pending_rows:
            self._show_warning("批量写回", "没有待写入的寄存器。")
            return
        confirm = QMessageBox.question(
            self,
            "确认写入",
            f"确认写入 {len(pending_rows)} 个已修改寄存器？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return

        unit_id = int(self.spin_unit.value())
        pending_rows.sort(key=lambda x: x.address)
        groups: List[List[RegisterRowState]] = []
        for row in pending_rows:
            if not groups or row.address != groups[-1][-1].address + 1:
                groups.append([row])
            else:
                groups[-1].append(row)

        touched_addrs = [row.address for row in pending_rows]
        for group in groups:
            try:
                if len(group) > 1:
                    start = group[0].address
                    values = [r.write_value for r in group]
                    self.append_log(
                        "TX",
                        f"批量写寄存器：address={start}, count={len(values)}, values={values}",
                    )
                    self._client.write_multiple_registers(unit_id, start, values)
                    self.append_log("OK", f"批量写入成功：start={start}, count={len(values)}")
                else:
                    row = group[0]
                    self.append_log("TX", f"写单寄存器：address={row.address}, value={row.write_value}")
                    self._client.write_single_register(unit_id, row.address, row.write_value)
                    self.append_log("OK", f"写入成功：address={row.address}")
            except Exception as exc:  # noqa: BLE001
                for row in group:
                    row.status = "写入失败"
                self.append_log("ERROR", f"写入失败：{self._format_user_exception(exc)}")

        try:
            verify_map = self._read_addresses_holding([row.address for row in self._table_rows])
            for row in self._table_rows:
                if row.address in verify_map:
                    row.current_value = verify_map[row.address]
                    self._last_values_by_addr[row.address] = row.current_value
                    if row.address in touched_addrs:
                        self.append_log(
                            "RX",
                            f"重读校验：address={row.address}, value={row.current_value}",
                        )
                        if row.current_value == row.write_value:
                            row.status = "写入成功"
                        else:
                            row.status = "写入失败"
                            self.append_log(
                                "ERROR",
                                f"写入失败：address={row.address}, expected={row.write_value}, actual={row.current_value}",
                            )
            self._updating_table = True
            for idx in range(len(self._table_rows)):
                self._refresh_table_row(idx)
            self._updating_table = False
            self._last_registers = [row.current_value for row in self._table_rows]
            self._apply_parsed_column()
            self._update_table_write_controls()
        except Exception as exc:  # noqa: BLE001
            self._show_error("重读校验失败", self._format_user_exception(exc))

    def _on_clear_result(self) -> None:
        self._clear_result_table()
        self.append_log("INFO", "已通过快捷按钮清空结果。")

    def _on_reset_config(self) -> None:
        self.edit_host.setText("127.0.0.1")
        self.spin_port.setValue(502)
        self.spin_baud.setValue(9600)
        self.spin_bytesize.setCurrentText("8")
        self.combo_parity.setCurrentText("N")
        self.combo_stopbits.setCurrentText("1")
        self.spin_unit.setValue(1)
        self.spin_addr.setValue(0)
        self.edit_batch_addrs.clear()
        self.spin_count.setValue(10)
        self.edit_values.clear()
        self.spin_poll_interval.setValue(1000)
        self.chk_poll.setChecked(False)
        self.combo_fc.setCurrentIndex(0)
        self.combo_dtype.setCurrentText("uint16")
        self.combo_endian.setCurrentText("AB CD")
        self._success_count = 0
        self._failure_count = 0
        self._timeout_count = 0
        self._refresh_stats_labels()
        self._clear_result_table()
        self.append_log("INFO", "已重置配置为默认值。")

    # ------------------------------------------------------------------ 配置 / 日志 / CSV
    def _gather_config(self) -> dict:
        return {
            "app": APP_NAME,
            "version": APP_VERSION,
            "comm_type": self.combo_conn.currentText(),
            "tcp_host": self.edit_host.text(),
            "tcp_port": int(self.spin_port.value()),
            "rtu_port": self._current_serial_device(),
            "baudrate": int(self.spin_baud.value()),
            "bytesize": int(self.spin_bytesize.currentText()),
            "parity": self.combo_parity.currentText(),
            "stopbits": float(self.combo_stopbits.currentText()),
            "unit_id": int(self.spin_unit.value()),
            "function_code": self._current_fc(),
            "start_address": int(self.spin_addr.value()),
            "batch_addresses": self.edit_batch_addrs.text().strip(),
            "count": int(self.spin_count.value()),
            "data_type": self.combo_dtype.currentText(),
            "word_order": self.combo_endian.currentText(),
            "poll_interval_ms": int(self.spin_poll_interval.value()),
            "poll_enabled": self.chk_poll.isChecked(),
        }

    def _apply_config(self, cfg: dict) -> None:
        ct = str(cfg.get("comm_type", "RTU"))
        self.combo_conn.blockSignals(True)
        self.combo_conn.setCurrentIndex(0 if ct == "TCP" else 1)
        self.combo_conn.blockSignals(False)
        self._tcp_params_page.setVisible(ct == "TCP")
        self._rtu_params_page.setVisible(ct != "TCP")
        self.edit_host.setText(str(cfg.get("tcp_host", "127.0.0.1")))
        self.spin_port.setValue(int(cfg.get("tcp_port", 502)))
        rp = str(cfg.get("rtu_port", ""))
        if rp:
            idx = self.combo_serial.findText(rp)
            if idx >= 0:
                self.combo_serial.setCurrentIndex(idx)
            else:
                self.combo_serial.setEditText(rp)
        self.spin_baud.setValue(int(cfg.get("baudrate", 9600)))
        bs = str(cfg.get("bytesize", "8"))
        i = self.spin_bytesize.findText(bs)
        if i >= 0:
            self.spin_bytesize.setCurrentIndex(i)
        py = str(cfg.get("parity", "N"))
        i = self.combo_parity.findText(py)
        if i >= 0:
            self.combo_parity.setCurrentIndex(i)
        st = str(cfg.get("stopbits", "1"))
        i = self.combo_stopbits.findText(st)
        if i >= 0:
            self.combo_stopbits.setCurrentIndex(i)
        self.spin_unit.setValue(int(cfg.get("unit_id", 1)))
        fc = str(cfg.get("function_code", "03"))
        for idx in range(self.combo_fc.count()):
            if str(self.combo_fc.itemData(idx)) == fc:
                self.combo_fc.setCurrentIndex(idx)
                break
        self.spin_addr.setValue(int(cfg.get("start_address", 0)))
        self.edit_batch_addrs.setText(str(cfg.get("batch_addresses", "")))
        self.spin_count.setValue(int(cfg.get("count", 10)))
        dt = str(cfg.get("data_type", "uint16"))
        i = self.combo_dtype.findText(dt)
        if i >= 0:
            self.combo_dtype.setCurrentIndex(i)
        wo = str(cfg.get("word_order", "AB CD"))
        i = self.combo_endian.findText(wo)
        if i >= 0:
            self.combo_endian.setCurrentIndex(i)
        self.spin_poll_interval.setValue(int(cfg.get("poll_interval_ms", 1000)))
        self.chk_poll.setChecked(bool(cfg.get("poll_enabled", False)))
        self._update_fc_dependent_widgets()
        self._update_poll_controls_enabled()

    def _on_save_config(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "保存配置",
            str(default_config_path()),
            "JSON (*.json)",
        )
        if not path:
            return
        try:
            save_config(Path(path), self._gather_config())
            self.append_log("OK", f"配置已保存：{path}")
        except Exception as exc:  # noqa: BLE001
            self._show_error("保存配置失败", str(exc))

    def _on_load_config(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "加载配置",
            str(default_config_path().parent),
            "JSON (*.json)",
        )
        if not path:
            return
        try:
            cfg = load_config(Path(path))
            self._refresh_serial_ports(log_info=False)
            self._apply_config(cfg)
            self.append_log("OK", f"配置已加载：{path}")
        except Exception as exc:  # noqa: BLE001
            self._show_error("加载配置失败", str(exc))

    def _on_clear_log(self) -> None:
        self.log.clear()
        self.append_log("INFO", "日志已清空。")

    def _on_save_log(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "保存日志",
            "",
            "文本文件 (*.txt)",
        )
        if not path:
            return
        try:
            Path(path).write_text(self.log.toPlainText(), encoding="utf-8")
            self.append_log("OK", f"日志已保存：{path}")
        except Exception as exc:  # noqa: BLE001
            self._show_error("保存日志失败", str(exc))

    def _on_export_csv(self) -> None:
        if self.table.rowCount() == 0:
            self.append_log("WARN", "暂无可导出的数据")
            QMessageBox.information(self, "导出 CSV", "暂无可导出的数据")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出 CSV",
            "",
            "CSV (*.csv)",
        )
        if not path:
            return
        try:
            p = Path(path)
            cols = [
                self.table.horizontalHeaderItem(c).text()
                for c in range(self.table.columnCount())
            ]
            with p.open("w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(cols)
                for r in range(self.table.rowCount()):
                    row = []
                    for c in range(self.table.columnCount()):
                        it = self.table.item(r, c)
                        row.append(it.text() if it else "")
                    w.writerow(row)
            self.append_log("OK", f"CSV 已导出：{path}")
        except Exception as exc:  # noqa: BLE001
            self._show_error("导出 CSV 失败", str(exc))

    def closeEvent(self, event) -> None:  # type: ignore[override]
        self._stop_polling_internal(silent=True)
        self._client.close()
        super().closeEvent(event)
