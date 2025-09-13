import time
import threading
import win32gui
import win32process
import win32con
import psutil
import os
import re

def find_process_tids(process_name):
    """查找指定进程名的所有线程ID"""
    tids = []
    process_name = process_name.lower()
    
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == process_name:
            try:
                # 获取进程的所有线程
                for thread in proc.threads():
                    tids.append((proc.info['pid'], thread.id))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
                
    return tids

def get_window_title(window_handle):
    """获取窗口标题"""
    if not win32gui.IsWindowVisible(window_handle):
        return None
        
    # 检查窗口样式
    style = win32gui.GetWindowLong(window_handle, win32con.GWL_STYLE)
    ex_style = win32gui.GetWindowLong(window_handle, win32con.GWL_EXSTYLE)
    
    # 排除子窗口和工具窗口
    if style & win32con.WS_CHILD or ex_style & win32con.WS_EX_TOOLWINDOW:
        return None
        
    # 获取窗口标题
    title = win32gui.GetWindowText(window_handle)
    return title if title else None

def extract_song_info(title):
    """从窗口标题中提取歌曲名和歌手名"""
    if not title:
        return "当前未播放歌曲"
    
    # 网易云音乐标题格式通常为: 歌曲名 - 歌手名 - 网易云音乐
    # QQ音乐标题格式通常为: 歌曲名 - 歌手名 - QQ音乐
    patterns = [
        r'^(.*?) - (.*?) - (网易云音乐|QQ音乐)$',
        r'^(.*?) - (.*?)$',
        r'^(.*?) — (.*?)$',
    ]
    
    for pattern in patterns:
        match = re.match(pattern, title)
        if match:
            song_name = match.group(1).strip()
            artist_name = match.group(2).strip()
            return f"{song_name} - {artist_name}"
    
    # 如果无法匹配模式，返回原始标题
    return title

def find_named_windows(tid):
    """通过线程ID查找所有命名窗口"""
    windows = {}
    
    def enum_thread_windows_callback(hwnd, windows_dict):
        title = get_window_title(hwnd)
        if title:
            windows_dict[hwnd] = title
        return True
    
    # 枚举线程的所有窗口
    win32gui.EnumThreadWindows(tid, enum_thread_windows_callback, windows)
    return windows

def write_to_file(content, filename="music.txt"):
    """将内容写入文件"""
    try:
        # 获取当前目录
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, filename)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content + '\n')
        print(f"已写入文件: {content}")
        return True
    except Exception as e:
        print(f"写入文件失败: {e}")
        return False

def find_player(process_name, player_name):
    """查找播放器进程和窗口"""
    while True:
        time.sleep(2)
        try:
            tids = find_process_tids(process_name)
            
            if not tids:
                # print(f"未找到{player_name}进程，继续搜索...")
                continue
                
            for pid, tid in tids:
                try:
                    windows = find_named_windows(tid)
                    if windows and len(windows) >= 1:
                        # 选择第一个窗口
                        window, title = list(windows.items())[0]
                        print(f"找到{player_name}窗口: {title}")
                        
                        # 提取歌曲信息并写入文件
                        song_info = extract_song_info(title)
                        write_to_file(song_info)
                        
                        # 启动监控线程
                        monitor_thread = threading.Thread(
                            target=update_player, 
                            args=(process_name, player_name, window)
                        )
                        monitor_thread.daemon = True
                        monitor_thread.start()
                        return
                except Exception as e:
                    print(f"处理{player_name}线程时出错: {e}")
                    continue
        except Exception as e:
            print(f"查找{player_name}时发生错误: {e}")
            time.sleep(5)

def update_player(process_name, player_name, window_handle):
    """监控播放器窗口标题变化"""
    last_title = None
    print(f"开始监控{player_name}窗口变化...")
    
    while True:
        time.sleep(1)
        
        try:
            # 检查窗口是否仍然存在
            if not win32gui.IsWindow(window_handle):
                print(f"{player_name} - 窗口已关闭")
                write_to_file("当前未播放歌曲")
                
                # 重新开始查找
                find_thread = threading.Thread(
                    target=find_player, 
                    args=(process_name, player_name)
                )
                find_thread.daemon = True
                find_thread.start()
                return
            
            # 获取当前标题
            current_title = get_window_title(window_handle)
            
            if current_title and current_title != last_title:
                last_title = current_title
                # 提取歌曲信息并写入文件
                song_info = extract_song_info(current_title)
                write_to_file(song_info)
                
        except Exception as e:
            print(f"监控{player_name}时出错: {e}")
            # 出错后重新开始查找
            find_thread = threading.Thread(
                target=find_player, 
                args=(process_name, player_name)
            )
            find_thread.daemon = True
            find_thread.start()
            return

def main():
    print("开始监控")
    
    # 创建文件
    script_dir = os.path.dirname(os.path.abspath(__file__))
    initial_file = os.path.join(script_dir, "music.txt")
    with open(initial_file, 'w', encoding='utf-8') as f:
        f.write("等待检测音乐播放器...\n")
    
    # 启动监控线程
    cloudmusic_thread = threading.Thread(
        target=find_player, 
        args=("cloudmusic.exe", "网易云音乐")
    )
    cloudmusic_thread.daemon = True
    cloudmusic_thread.start()
    
    qqmusic_thread = threading.Thread(
        target=find_player, 
        args=("qqmusic.exe", "QQ音乐")
    )
    qqmusic_thread.daemon = True
    qqmusic_thread.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n程序已停止")
        write_to_file("监控未开启")

if __name__ == "__main__":
    main()