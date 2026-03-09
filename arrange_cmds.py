import win32gui
import win32con
import win32api
import math
import time
import ctypes
import os
from ctypes import wintypes

# --- 配置 ---
USER32 = ctypes.windll.user32
KERNEL32 = ctypes.windll.kernel32

def minimize_self_immediately():
    """尝试立即最小化当前脚本所在的控制台窗口，并返回其句柄。"""
    current_pid = os.getpid()
    hwnd = None
    
    # 方法 1: GetConsoleWindow (适用于独立 cmd.exe)
    try:
        hwnd = KERNEL32.GetConsoleWindow()
        if hwnd:
            pid = wintypes.DWORD()
            USER32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == current_pid:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                return hwnd 
    except Exception:
        pass

    # 方法 2: 枚举查找 (备用)
    target_classes = ["ConsoleWindowClass", "CASCADIA_HOSTING_WINDOW_CLASS"]
    found_hwnds = []
    
    def enum_callback(hwnd, results):
        if not win32gui.IsWindowVisible(hwnd): return True
        try:
            pid = wintypes.DWORD()
            USER32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == current_pid:
                class_name = win32gui.GetClassName(hwnd)
                if class_name in target_classes:
                    results.append(hwnd)
        except: pass
        return True

    win32gui.EnumWindows(enum_callback, found_hwnds)
    
    if found_hwnds:
        h = found_hwnds[0]
        win32gui.ShowWindow(h, win32con.SW_MINIMIZE)
        return h
    
    return None

def get_cmd_windows(exclude_hwnd=None):
    """获取所有可见的 CMD/PowerShell 窗口，排除指定句柄。"""
    hwnds = []
    def enum_callback(hwnd, results):
        if not win32gui.IsWindowVisible(hwnd): return True
        try:
            class_name = win32gui.GetClassName(hwnd)
            if class_name in ["ConsoleWindowClass", "CASCADIA_HOSTING_WINDOW_CLASS"]:
                if exclude_hwnd and hwnd == exclude_hwnd: return True
                if win32gui.GetWindowText(hwnd): # 确保有标题
                    results.append(hwnd)
        except: pass
        return True
    win32gui.EnumWindows(enum_callback, hwnds)
    return hwnds

def minimize_other_windows(protected_hwnds):
    """最小化所有非 CMD 且非保护的窗口。"""
    count = 0
    protected_set = set(protected_hwnds) if protected_hwnds else set()
    system_classes = {"Shell_TrayWnd", "Progman", "WorkerW", "IME", "MSCTFIME UI", "ApplicationFrameHost"}

    def enum_callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd): return True
        try:
            if hwnd in protected_set: return True
            class_name = win32gui.GetClassName(hwnd)
            if class_name in system_classes: return True
            # 如果是 CMD 窗口，保留（留给排列函数处理）
            if class_name in ["ConsoleWindowClass", "CASCADIA_HOSTING_WINDOW_CLASS"]: return True
            
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            nonlocal count
            count += 1
        except: pass
        return True

    win32gui.EnumWindows(enum_callback, None)
    return count

def get_work_area():
    rect = ctypes.wintypes.RECT()
    if USER32.SystemParametersInfoA(48, 0, ctypes.byref(rect), 0):
        return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
    return 0, 0, win32api.GetSystemMetrics(win32con.SM_CXSCREEN), win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

def arrange_windows_grid(hwnds, max_rows=5):
    if not hwnds:
        print("没有需要排列的 CMD 窗口。")
        return

    count = len(hwnds)
    # 注意：这里的 count 是包含第一个窗口的总数，但实际排列数会少 1
    actual_count = count - 1 if count > 0 else 0
    
    if actual_count <= 0:
        print("排除第一个窗口后，没有剩余窗口需要排列。")
        return

    print(f"准备排列 {actual_count} 个 CMD 窗口 (已排除第 1 个)。")

    # --- 智能计算行数 (基于实际排列数量) ---
    # 使用 actual_count 来计算网格，避免因为排除了一个窗口导致网格过大或过小
    base_sqrt = math.sqrt(actual_count)
    rows = math.ceil(base_sqrt)
    if rows > max_rows:
        rows = max_rows
    
    # 优化行数逻辑 (同之前版本)
    if rows > 1:
        if actual_count <= rows * (rows - 1):
            rows -= 1
    if rows > max_rows:
        rows = max_rows
        
    cols = math.ceil(actual_count / rows)
    
    # 防止列数为 0 (虽然 actual_count > 0 保证了这点)
    if cols == 0: cols = 1

    print(f"智能布局计算：{actual_count} 个窗口 -> {rows} 行 x {cols} 列")

    # 获取工作区
    try:
        start_x, start_y, work_width, work_height = get_work_area_ctypes()
    except Exception as e:
        print(f"获取工作区失败: {e}")
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

    # --- 核心修改：遍历并跳过第一个 ---
    for index, hwnd in enumerate(hwnds):
        # 【新增】如果是第一个窗口 (index 0)，直接跳过，不排列
        if index == 0:
            # 可选：如果你想让被跳过的窗口保持最小化，可以在这里加一行代码
            # win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            continue 

        # 计算当前是第几个被排列的窗口 (用于日志和坐标计算)
        # 因为跳过了 index 0，所以当前的逻辑序号应该是 index - 1
        current_arrange_index = index - 1
        
        row = current_arrange_index // cols
        col = current_arrange_index % cols
        
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
            
            # 日志输出优化：基于 current_arrange_index 计数
            if current_arrange_index < 4 or current_arrange_index == actual_count - 1:
                print(f"  [{current_arrange_index+1}] {title[:25]}...")
            elif current_arrange_index == 4:
                print(f"  ... (剩余 {actual_count-5} 个窗口已排列)")
                
        except Exception as e:
            print(f"  [{current_arrange_index+1}] 失败: {e}")

def main():
    # 1. 最小化自己
    self_hwnd = minimize_self_immediately()
    time.sleep(0.2)

    # 2. 获取其他 CMD
    targets = get_cmd_windows(exclude_hwnd=self_hwnd)
    if not targets:
        print("未找到其他 CMD 窗口。")
        return

    # 3. 清理桌面
    cnt = minimize_other_windows(targets)
    if cnt > 0:
        print(f"已最小化 {cnt} 个无关窗口。")

    # 4. 排列
    print(f"正在排列 {len(targets)} 个 CMD 窗口...")
    arrange_windows_grid(targets, max_rows=5)
    print("完成！")

if __name__ == "__main__":
    try:
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("提示：建议以管理员身份运行。")
    except: pass
    main()
