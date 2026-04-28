# Modbus Studio V3

基于 **Python + PySide6 + pymodbus** 的桌面版 Modbus 调试工具，在 V2 基础上增强了批量地址监控、变化高亮、自动重连、请求节流、错误统计与快捷操作能力。

## 版本与常量

- 应用名：`Modbus Studio`（代码中 `APP_NAME`）
- 当前版本：`0.3.0`（`modbus_tool/version.py` 中 `APP_VERSION`）

## 功能概览

### 通讯

- **默认 RTU**：启动后通讯类型为 RTU，显示串口参数；TCP 保留，手动切换后显示 IP/端口。
- **Modbus TCP / RTU**：连接、断开、从站地址（Unit ID，支持 `1~254`，即 `0x01~0xFE`）。
- **串口**：启动自动扫描；**刷新串口**重新枚举；下拉仅显示设备名（如 `COM3`）。无可用串口时日志 `[WARN] 未检测到可用串口`。

### 读写与轮询

- **功能码**：01 读线圈、02 读离散输入、03 读保持寄存器、04 读输入寄存器、05 写单线圈、06 写单寄存器、0F 写多线圈、16 写多寄存器。
- **周期轮询**（仅 01/02/03/04）：启用轮询、间隔（默认 1000 ms，最小 300 ms）、开始/停止；与手动执行互斥（未完成则跳过本次定时器）；断开连接自动停止。

### 数据解析

- **类型**：uint16、int16、uint32、int32、float32（默认 uint16）。
- **字序**：AB CD（高字在前）、CD AB（字交换）；32 位与 float 每 **2 个寄存器** 一组，不足 2 个时解析列为空。
- **结果表**：地址、原始十进制、原始十六进制、二进制、解析值。

### 配置与导出

- **保存配置 / 加载配置**：JSON，默认建议路径 `config/modbus_tool_config.json`（相对项目根目录，目录不存在会自动创建）。
- **日志**：清空日志、保存日志为 `.txt`（自选路径）。
- **导出 CSV**：导出当前表格全部列。

## 安装依赖

在项目根目录（与 `main.py` 同级）执行：

```bash
pip install -r requirements.txt
```

依赖：`PySide6`、`pymodbus`、`pyserial`。

## 启动方式

```bash
python main.py
```

请在项目根目录运行，以便正确加载包 `modbus_tool` 与默认 `config/` 路径。

## 项目结构（节选）

```text
main.py
requirements.txt
README.md
config/                         # 可选：保存 modbus_tool_config.json
modbus_tool/
  __init__.py
  version.py                    # APP_NAME / APP_VERSION
  app.py
  core/
    __init__.py
    modbus_client.py
    serial_utils.py
    data_parser.py
    config_manager.py
  ui/
    main_window.py
```

## 说明

- 不包含数据库、插件系统或复杂主题框架。
- 轮询使用 **QTimer**，在同线程内执行；长时间阻塞仍可能影响界面响应。
