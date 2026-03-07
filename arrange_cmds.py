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

def arrange_windows_grid(hwnds, max_rows=5):
    """
    将窗口排列成网格。
    优化逻辑：动态计算行数，优先铺满垂直空间，避免窗口过小。
    :param hwnds: 窗口句柄列表
    :param max_rows: 允许的最大行数，防止窗口变得太扁
    """
    if not hwnds:
        print("没有需要排列的 CMD 窗口。")
        return

    count = len(hwnds)
    print(f"准备排列 {count} 个 CMD 窗口。")

    # --- [核心优化] 动态计算行数 (Rows) 和列数 (Cols) ---
    
    # 策略：
    # 1. 计算理想的正方形边长 sqrt(count)
    # 2. 尝试减少行数，让列数增加，从而增加每个窗口的高度，直到行数达到下限或列数过多
    # 简单策略：根据数量区间固定行数，但允许在区间内动态调整以铺满
    
    # 基础行数估算 (向下取整的平方根，或者根据区间)
    # 如果 count=10, sqrt=3.16. 
    # 方案 A (原逻辑): 4x4 (浪费空间)
    # 方案 B (新逻辑): 3 行 -> ceil(10/3)=4 列 (3x4=12 格子，只空 2 格，高度增加)
    
    base_sqrt = math.sqrt(count)
    
    # 确定行数：
    # 如果数量很少，至少 1 行。
    # 我们希望行数尽可能少（以增加高度），但不能让列数无限多（导致窗口太窄）。
    # 限制：行数 <= max_rows, 且 行数 >= 1
    # 算法：从 ceil(sqrt(count)) 开始尝试减少行数，直到 (count / rows) 的列宽看起来合理？
    # 更简单的启发式算法：
    # 如果 count <= 4: 1 行 or 2 行? -> 2 行比较稳，除非只有 1 个
    # 如果 count <= 9: 3 行
    # 如果 count <= 16: 4 行
    # 如果 count <= 25: 5 行
    # 修正：如果 count=10，按上面是 4 行。但我们希望它是 3 行。
    # 规则：如果 count > (rows-1)^2 且 count <= rows^2:
    #       如果 count 接近 (rows-1)*rows，则使用 rows-1 行。
    
    rows = math.ceil(base_sqrt)
    if rows > max_rows:
        rows = max_rows
    
    # 优化检查：如果当前行数下，最后一行空的太多，尝试减少一行
    # 例如：10 个窗口，rows=4 (4x4=16, 空 6 个). 
    # 尝试 rows=3 (3x4=12, 空 2 个). 显然 3 行更好。
    # 条件：如果 (rows * (rows-1)) >= count，说明减少一行也能装下，且利用率更高（窗口更高）
    if rows > 1:
        if (rows - 1) * math.ceil(count / (rows - 1)) >= count: 
            # 这里逻辑稍微修正：只要 (rows-1) 行能装下（即 cols 增加一点但总格子够），就减少行数
            # 实际上只要 count <= (rows-1) * some_col，我们想最大化高度。
            # 最简单的判断：如果 count <= rows * (rows - 1)，那么用 rows-1 行会更紧凑且更高。
            # 例：10 <= 4*3 (12)? 是的。所以用 3 行。
            # 例：13 <= 4*3 (12)? 否。必须用 4 行。
            if count <= rows * (rows - 1):
                rows -= 1

    # 确保不超过最大限制
    if rows > max_rows:
        rows = max_rows
        
    # 计算列数
    cols = math.ceil(count / rows)

    print(f"智能布局计算：{count} 个窗口 -> {rows} 行 x {cols} 列 (最大化窗口高度)")

    # 获取工作区
    try:
        start_x, start_y, work_width, work_height = get_work_area_ctypes()
    except Exception as e:
        print(f"获取工作区失败: {e}, 使用全屏。")
        start_x, start_y = 0, 0
        work_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        work_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

    margin = 10
    available_width = work_width - (margin * (cols + 1))
    available_height = work_height - (margin * (rows + 1))

    if available_width <= 0 or available_height <= 0:
        print("错误：屏幕空间不足。")
        return

    cell_width = available_width // cols
    cell_height = available_height // rows

    print(f"单个窗口尺寸：{cell_width}x{cell_height}")

    for index, hwnd in enumerate(hwnds):
        row = index // cols
        col = index % cols
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
            # 仅打印前几个窗口的详细信息，避免刷屏
            if index < 5 or index == count - 1:
                print(f"  [{index+1}] {title[:20]}...")
            elif index == 5:
                print(f"  ... (剩余 {count-6} 个窗口已排列)")
                
        except Exception as e:
            print(f"  [{index+1}] 失败: {e}")

def main():
    print("="*30)
    print("CmdGrid - 智能窗口排列")
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

    # 3. 暂时最小化脚本窗口
    script_was_minimized = False
    if script_hwnd:
        print("正在暂时最小化脚本窗口...")
        win32gui.ShowWindow(script_hwnd, win32con.SW_MINIMIZE)
        script_was_minimized = True
        time.sleep(0.3)

    try:
        # 4. 构建保护列表
        protected_list = target_windows.copy()
        
        # 5. 最小化其他所有无关窗口
        minimize_other_windows(protected_list)

        # 6. 排列目标 CMD 窗口 (应用新的智能网格逻辑)
        arrange_windows_grid(target_windows, max_rows=5)
        
    finally:
        # 7. 还原脚本窗口
        if script_was_minimized and script_hwnd:
            print("\n正在还原脚本窗口...")
            time.sleep(0.5)
            win32gui.ShowWindow(script_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(script_hwnd)
            print("脚本窗口已还原。")

    print("="*30)
    print("完成！")
    print("="*30)

if __name__ == "__main__":
    try:
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("提示：建议以管理员身份运行。")
    except:
        pass
        
    main()
