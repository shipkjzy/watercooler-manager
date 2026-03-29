# Third-Party Notices

本项目包含或依赖以下第三方项目、代码、设计思路或二进制组件。它们分别受各自许可证约束。

---

## 1. tomups/watercooler-manager

- 项目名称：Water Cooler Manager
- 上游仓库：`tomups/watercooler-manager`
- 用途：提供水冷控制器 BLE 协议支持、设备控制逻辑与基础实现参考
- 许可证：GPL-3.0

说明：
本项目的部分设备支持能力、控制协议思路与基础行为来源于该上游项目。

---

## 2. noteMASTER11/watercooler-manager

- 项目名称：Watercooler Manager GUI
- 上游仓库：`noteMASTER11/watercooler-manager`
- 用途：提供 GUI 化实现基础、界面结构与集成方案参考
- 说明：该项目标明其基于 `tomups/watercooler-manager` 修改而来

说明：
本项目在其基础上继续进行了二次开发，包括但不限于：

- 配置管理增强
- 自动风扇/水泵曲线扩展
- 温控 RGB
- 管理员权限自动申请
- 主题与界面增强
- 防抖、回差与独立启停间隔

---

## 3. LibreHardwareMonitor

- 项目名称：LibreHardwareMonitor
- 上游仓库：`LibreHardwareMonitor/LibreHardwareMonitor`
- 用途：读取 CPU / GPU 等硬件温度与传感器数据
- 许可证：MPL-2.0

说明：
本项目通过 `LibreHardwareMonitorLib.dll` 访问硬件监控数据。
该组件及其相关文件仍受其原始许可证约束。

---

## 4. Python Dependencies

本项目还依赖若干 Python 库，包括但不限于：

- PyQt5
- qasync
- bleak
- pythonnet
- PyInstaller（构建时）

这些依赖分别受其各自许可证约束。分发本项目时，应同时遵守这些依赖的许可证要求。

---

## Redistribution Notes

当你分发本项目源码或二进制版本时，建议同时包含：

- 本项目的 `LICENSE`
- 本文件 `THIRD_PARTY_NOTICES.md`
- 与发布版本对应的完整源码
- 发布版本的变更说明

如果你分发了 `LibreHardwareMonitorLib.dll`，还应确保其许可证与相关说明一并提供。
