# Mouse Wheel Manager (Windows) 🖱️

一个基于 Python 的现代 GUI 工具，用于在 Windows 上单独管理鼠标滚轮方向（实现类似 macOS 的“自然滚动”）。

与其他修改注册表的方法不同，**本项目通过调用底层 Windows CfgMgr32 API 实现设备驱动的“软重启”**，修改设置后**无需物理插拔鼠标**即可立即生效。

## ✨ 功能特点

* **无需插拔 (Hot-Reload)**: 利用 `CM_Disable_DevNode` 和 `CM_Enable_DevNode` 自动重启目标设备，配置即刻生效。
* **智能识别**: 自动扫描 HID 和 Bluetooth 总线，通过 PID/VID 过滤并显示真实的设备名称（如 "Logitech MX Master 3S"）。
* **可视化管理**: 基于 `CustomTkinter` 的现代化 UI，支持浅色/深色模式自适应。
* **安全可靠**: 仅修改目标设备的 `FlipFlopWheel` 注册表项，不影响系统其他设置。

## 🛠️ 技术栈

* **语言**: Python 3.x
* **UI 框架**: [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
* **核心 API**:
    * `winreg`: 读写 Windows 注册表。
    * `ctypes (CfgMgr32)`: Windows 配置管理器 API，用于设备树管理和状态控制。

## 📦 安装与依赖

1. 克隆仓库：
   ```bash
   git clone ******
   cd MouseWheelManager
安装依赖：


pip install -r requirements.txt
🚀 使用方法
以管理员身份运行脚本（程序会自动请求 UAC 权限，因为涉及注册表和驱动操作）：

Bash

python main.py
(注：假设你的主代码文件名为 main.py)

在左侧列表中选择你的鼠标设备。

点击右侧的 "Mac 模式" 或 "Windows 模式"。

程序会自动写入注册表并重启该鼠标驱动。

⚠️ 注意事项
程序需要管理员权限才能运行。

部分只有基础驱动的鼠标（Generic HID Device）可能没有 FlipFlopWheel 属性，程序会自动过滤这些设备。