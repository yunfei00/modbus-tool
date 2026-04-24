"""应用入口：创建 QApplication 与主窗口。"""

import sys

from PySide6.QtWidgets import QApplication

from modbus_tool.ui.main_window import MainWindow


def main() -> None:
    """启动 GUI 事件循环。"""
    app = QApplication(sys.argv)
    app.setApplicationName("Modbus Tool")
    app.setOrganizationName("modbus-tool")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
