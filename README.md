# CmdGrid 🖥️✨

**English** | [中文](#中文说明)

## 🇬🇧 English Description

**CmdGrid** is a lightweight Python utility designed to automatically organize your scattered Command Prompt (CMD) windows into a neat grid layout. It cleans up your desktop by minimizing all unrelated applications and ensures your terminal workspace is tidy and efficient.

### ✨ Key Features

*   **🧹 Desktop Cleanup**: Automatically minimizes all non-CMD windows (browsers, folders, etc.) to reduce distractions.
*   **📐 Smart Grid Layout**: Arranges all open CMD windows into a perfect $N \times N$ grid based on the number of windows detected.
*   **👻 Stealth Mode**: The script window itself minimizes during execution to avoid blocking your view, then automatically restores upon completion to show results.
*   **🛡️ Self-Aware**: Intelligently identifies and excludes its own running instance from being moved or minimized permanently.
*   **⚡ Fast & Lightweight**: Built with `pywin32`, no heavy dependencies required.

### 📋 Prerequisites

*   **Python 3.x**
*   **Windows OS** (This tool uses Windows-specific APIs)
*   **Dependencies**:
    ```bash
    pip install pywin32
    ```

### 🚀 Usage

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/firzen/CmdGrid.git
    cd CmdGrid
    ```

2.  **Run the script**:
    *   **Recommended**: Run as **Administrator** to ensure it can minimize all system windows successfully.
    ```bash
    python arrange_cmds.py
    ```
    *(Or simply double-click `arrange_cmds.py` if you have Python associated, but Admin rights are preferred.)*

3.  **What happens next?**
    *   The script window briefly disappears.
    *   All other apps (Chrome, Explorer, etc.) minimize to the taskbar.
    *   Your CMD windows snap into a grid.
    *   The script window reappears with a "Done!" message.

### ⚙️ How it Works

1.  Scans for all visible windows with the class `ConsoleWindowClass`.
2.  Identifies its own Process ID (PID) to exclude itself.
3.  Minimizes the script window temporarily.
4.  Iterates through all other visible windows and minimizes them (excluding CMDs and system trays).
5.  Calculates the optimal grid size ($3\times3$, $4\times4$, etc.) based on the count of CMD windows.
6.  Restores and resizes each CMD window into the calculated grid positions.
7.  Restores the script window.

### 📄 License

MIT License

---

## 🇨🇳 中文说明

**CmdGrid** 是一个轻量级的 Python 实用工具，旨在将您散乱的命令提示符 (CMD) 窗口自动整理成整齐的网格布局。它通过最小化所有无关应用程序来清理桌面，确保您的终端工作区整洁高效。

### ✨ 主要功能

*   **🧹 桌面清理**: 自动最小化所有非 CMD 窗口（如浏览器、文件夹等），减少干扰。
*   **📐 智能网格布局**: 根据检测到的 CMD 窗口数量，自动将它们排列成完美的 $N \times N$ 网格。
*   **👻 隐身模式**: 脚本运行时会自动最小化自身窗口，避免遮挡视线；执行完毕后自动还原以显示结果。
*   **🛡️ 自我识别**: 智能识别并排除正在运行脚本的自身窗口，防止其被误操作或永久最小化。
*   **⚡ 快速轻量**: 基于 `pywin32` 构建，无需沉重的依赖库。

### 📋 环境要求

*   **Python 3.x**
*   **Windows 操作系统** (本工具使用 Windows 特有 API)
*   **依赖安装**:
    ```bash
    pip install pywin32
    ```

### 🚀 使用方法

1.  **克隆仓库**:
    ```bash
    git clone https://github.com/firzen/CmdGrid.git
    cd CmdGrid
    ```

2.  **运行脚本**:
    *   **推荐**: 以**管理员身份**运行，以确保能成功最小化所有系统窗口。
    ```bash
    python arrange_cmds.py
    ```
    *(或者直接双击 `arrange_cmds.py`，但建议赋予管理员权限以获得最佳效果。)*

3.  **运行效果**:
    *   脚本窗口会短暂消失。
    *   所有其他应用（Chrome、资源管理器等）最小化到任务栏。
    *   您的 CMD 窗口自动吸附到网格位置。
    *   脚本窗口重新弹出，显示“完成！”提示。

### ⚙️ 工作原理

1.  扫描所有类名为 `ConsoleWindowClass` 的可见窗口。
2.  识别自身的进程 ID (PID) 以排除自己。
3.  暂时最小化脚本窗口。
4.  遍历所有其他可见窗口并将其最小化（排除 CMD 窗口和系统托盘等关键窗口）。
5.  根据 CMD 窗口数量计算最佳网格大小（$3\times3$, $4\times4$ 等）。
6.  还原并调整每个 CMD 窗口的大小，将其放置在计算好的网格坐标上。
7.  还原脚本窗口。

### 📄 许可证

MIT License

---

### 💡 Tips / 提示


*   **Multiple Monitors / 多显示器**: Currently supports the primary monitor's work area.
    目前主要支持主显示器的工作区域。
