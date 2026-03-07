import win32gui
import win32con
import win32api
import math
import time
import ctypes
import os
from ctypes import wintypes

# 定义 SPI_GETWORKAREA 常量
SPI_GETWORKAREA = 48

def get_current_console_hwnd():
    """获取当前脚本运行的控制台窗口句柄。"""
    current_pid = os.getpid()
    found_hwnd = None
    
    def enum_callback(hwnd, _):
        nonlocal found_hwnd
        try:
            if win32gui.IsWindowVisible(hwnd):
                _, window_pid = win32api.GetWindowThreadProcessId(hwnd)
                if window_pid == current_pid:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "ConsoleWindowClass":
                        found_hwnd = hwnd
                        return False
        except Exception:
            pass
        return True

    win32gui.EnumWindows(enum_callback, None)
    return found_hwnd

def get_cmd_windows(exclude_hwnd=None):
    """获取所有可见的 CMD 窗口句柄列表（可排除指定句柄）。"""
    hwnds = []
    
    def enum_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            try:
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "ConsoleWindowClass":
                    if exclude_hwnd and hwnd == exclude_hwnd:
                        return True
                    results.append(hwnd)
            except Exception:
                pass
        return True

    win32gui.EnumWindows(enum_callback, hwnds)
    return hwnds

def minimize_other_windows(protected_hwnds):
    """最小化所有不在保护列表中的可见窗口。"""
    print("正在清理桌面：最小化无关窗口...")
    count = 0
    protected_set = set(protected_hwnds)
    
    system_classes = {"Shell_TrayWnd", "Progman", "WorkerW", "IME", "MSCTFIME UI"}

    def enum_callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            if hwnd in protected_set:
                return True
            class_name = win32gui.GetClassName(hwnd)
            if class_name in system_classes:
                return True
            
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            nonlocal count
            count += 1
        except Exception:
            pass
        return True

    win32gui.EnumWindows(enum_callback, None)
    print(f"已最小化 {count} 个无关窗口。")
    time.sleep(0.5) 

def get_work_area_ctypes():
    """使用 ctypes 获取屏幕工作区。"""
    rect = ctypes.wintypes.RECT()
    result = ctypes.windll.user32.SystemParametersInfoA(
        SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
    )
    if not result:
        raise RuntimeError("调用 SystemParametersInfo 失败")
    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top

def arrange_windows_grid(hwnds, grid_size=None):
    if not hwnds:
        print("没有需要排列的 CMD 窗口。")
        return

    count = len(hwnds)
    print(f"准备排列 {count} 个 CMD 窗口。")

    # 计算网格
    if grid_size is None:
        n = 3 if count <= 9 else (4 if count <= 16 else (5 if count <= 25 else math.ceil(math.sqrt(count))))
    else:
        n = grid_size
    
    if count > n*n:
        print(f"提示：窗口数 ({count}) 超过网格容量 ({n}x{n})，部分窗口将重叠或需手动调整。")

    # 获取工作区
    try:
        start_x, start_y, work_width, work_height = get_work_area_ctypes()
    except Exception as e:
        print(f"获取工作区失败: {e}, 使用全屏。")
        start_x, start_y = 0, 0
        work_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        work_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    margin = 10
    available_width = work_width - (margin * (n + 1))
    available_height = work_height - (margin * (n + 1))

    if available_width <= 0 or available_height <= 0:
        print("错误：屏幕空间不足。")
        return

    cell_width = available_width // n
    cell_height = available_height // n

    print(f"布局：{n}x{n}, 窗口尺寸：{cell_width}x{cell_height}")

    for index, hwnd in enumerate(hwnds):
        row = index // n
        col = index % n
        x = start_x + margin + (col * cell_width)
        y = start_y + margin + (row * cell_height)
        
        try:
            # 还原并置顶
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.02)
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOP, x, y, cell_width, cell_height,
                win32con.SWP_NOZORDER | win32con.SWP_SHOWWINDOW
            )
            title = win32gui.GetWindowText(hwnd)
            print(f"  [{index+1}] {title[:20]}...")
        except Exception as e:
            print(f"  [{index+1}] 失败: {e}")

def main():
    print("="*30)
    print("CMD 窗口自动排列工具")
    print("="*30)
    
    # 1. 获取脚本自身窗口
    script_hwnd = get_current_console_hwnd()
    
    # 2. 获取其他 CMD 窗口 (排除脚本窗口)
    target_windows = get_cmd_windows(exclude_hwnd=script_hwnd)
    
    if not target_windows:
        msg = "未找到其他 CMD 窗口。"
        if script_hwnd:
            msg += "\n(当前只有脚本窗口自己)"
        print(msg)
        return

    # 3. 【关键步骤】如果脚本窗口存在，先将其最小化，以免遮挡视线
    script_was_minimized = False
    if script_hwnd:
        print("正在暂时最小化脚本窗口...")
        # 检查当前状态，如果是最小化就不需要再操作，但为了保险直接发最小化指令
        win32gui.ShowWindow(script_hwnd, win32con.SW_MINIMIZE)
        script_was_minimized = True
        time.sleep(0.3) # 等待最小化动画完成

    try:
        # 4. 构建保护列表 (只保护目标 CMD，脚本窗口此时已最小化，不需要在保护列表里防止被最小化)
        # 但为了防止逻辑混乱，我们依然把 target_windows 作为保护对象
        protected_list = target_windows.copy()
        
        # 5. 最小化其他所有无关窗口
        minimize_other_windows(protected_list)

        # 6. 排列目标 CMD 窗口
        arrange_windows_grid(target_windows)
        
    finally:
        # 7. 【关键步骤】无论成功失败，最后都要还原脚本窗口
        if script_was_minimized and script_hwnd:
            print("\n正在还原脚本窗口...")
            time.sleep(0.5) # 等排列稍微稳定一点
            win32gui.ShowWindow(script_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(script_hwnd) # 尝试将焦点给回脚本窗口
            print("脚本窗口已还原。")

    print("="*30)
    print("完成！")
    print("="*30)

if __name__ == "__main__":
    # 检查管理员权限提示
    try:
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("提示：建议以管理员身份运行，以便最小化高权限窗口。")
    except:
        pass
        
    main()
