import pandas
import flask
import json
import pathlib
import psutil
import os
import yaml
import time
import threading
import sys
import logging
import subprocess
import win32gui,win32print,win32con
import win32process
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt,QObject, Signal,QRect
from PySide6.QtUiTools import QUiLoader
import tqdm

DISKS = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","S","U","V","W","X","Y","Z"]
disks = []
file_count = 10000
FlagFileCreated = False
bar = tqdm.tqdm(total=10000, desc="文件已扫描", ncols=100, ascii=True)
ExecutebaleName = ""
ProcessName = ""
xlsx_path = ""
server = flask.Flask(__name__)

path = None
w = None
t = 1
#log
def setup_logger():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    file_handler = logging.FileHandler('./log.log',encoding="utf-8",mode="w")
    file_handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    return logger
logger = setup_logger()
logger.info("已初始化日志系统")
logger.info("程序启动中...")

def get_cache_path():
    global xlsx_path
    if getattr(sys, 'frozen', False):
        BASE_PATH = os.path.dirname(sys.executable)
        xlsx_file = pathlib.Path(BASE_PATH).joinpath("translation_data")
        xlsx_file.mkdir(exist_ok=True)
        appdata_dir = pathlib.Path(os.environ['APPDATA'])
        cache_dir = appdata_dir / "FeatherTranslator"
        cache_dir.mkdir(exist_ok=True)
        xlsx_path = xlsx_file / "translates.xlsx"
        cache_file = cache_dir / "gamepath.yaml"       
        return cache_file
    else:
        cache_file = pathlib.Path(__file__).parent / "gamepath.yaml"
        return cache_file
    
# 创建信号类用于线程间通信
class OverlaySignals(QObject):
    update_text = Signal(str)

class Translate:
    def __init__(self):
        try:
            if getattr(sys, 'frozen', False):
                file = xlsx_path
            else:
                path = pathlib.Path(__file__).resolve().parent
                file = path / "translates.xlsx"
            self.file = pandas.read_excel(file,sheet_name="translate")
            self.file_config = pandas.read_excel(file,sheet_name="IdentityData")
            logger.info("已加载翻译文件")
        except Exception as e:
            logger.error(f"无法读取翻译文件:{e}")

    def get(self, name:str):
        for index in self.file.index:
            if name == self.file.at[index, "音频"]:
                logger.info(f"已找到{self.file.at[index, '音频']}对应的翻译文本:{self.file.at[index, '翻译']}")
                return self.file.at[index, "翻译"]
trans = Translate()
config = trans.file_config

ExecutebaleName = config.at[0,"values"]
ProcessName = config.at[1,"values"]

def process(raw:dict, to_cli:bool = True):
    name:str = raw["name"]
    out = trans.get(name)
    if to_cli:
        print(out)
        return None
    else:
        return out 

def extract_zip(path) -> None:
    source_path = "./Beplnex.zip"
    args = f"-Path {source_path} -DestinationPath {path}"
    cmd = f"powershell.exe -Command Expand-Archive {args}"
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    logger.info(f"powershell退出code:{result.returncode}")
    logger.info(f"powershell输出:{result.stdout.strip()}")

def check_alive_and_move():
    time.sleep(0.2)
    window_move()
    process_name = ProcessName
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == process_name.lower():
            return True
    return False

def check_alive():
    time.sleep(0.2)
    process_name = ProcessName
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == process_name.lower():
            return True
    return False

def get_size():
    hDC = win32gui.GetDC(0)
    w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    logger.info(f"已获取屏幕分辨率:{w}x{h}")
    return w, h

def check_file(paths):
    disk_list = paths
    logger.info(f"正在尝试全局扫描游戏主程序...")
    for i in range(len(paths)):
        path = disk_list[i]
        logger.info(f"当前查找磁盘:{path}")
        count = 0
        for filepath,dirname,filename in os.walk(path):
            if "windows" in str(filepath.lower()):
                pass
            else:
                for temp in filename:
                    count += 1
                    bar.update(1)
                    if count >= file_count:
                        count = 0
                    if temp == ExecutebaleName:
                        bar.clear()
                        full_path = os.path.join(filepath,temp)
                        return full_path
                    else:
                        pass
    return ""

def extract_plugin(path) -> None:
    source_path = "./Plugin.zip"
    if not os.path.exists(os.path.join(path,"BepInEx/plugins")):
        logger.error("无法找到plugin文件夹,将自动启动游戏初始化ModLoader,请在游戏窗口消失后重启程序。")
        exe_dir = os.path.join(path,ExecutebaleName)
        os.startfile(exe_dir)
        time.sleep(5)
        if check_alive():
            for proc in psutil.process_iter(['pid']):
                if proc.info['pid']:
                    pid = proc.info['pid']
                    psutil.Process(pid).kill()
    Path = os.path.join(path,"BepInEx/plugins")
    args = f"-Path {source_path} -DestinationPath {Path}"
    cmd = f"powershell.exe -Command Expand-Archive {args}"
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    logger.info(f"powershell退出code:{result.returncode}")
    logger.info(f"powershell输出:{result.stdout.strip()}")
        
def locate_game():
    for disk in DISKS:
        disk = disk+":/"
        path = pathlib.Path(disk)
        if path.exists():
            disks.append(path)
        else:
            pass
    logger.info(f"存在磁盘:{disks}")
    full_path = check_file(disks)
    if full_path != "":
        logger.info(f"已定位到游戏，路径:{full_path}")
        return full_path
    else:
        return "Not found"

def pull_up_game():
    global w
    global path
    w,h = get_size()
    path = pathlib.Path(__file__).resolve().parent
    file = get_cache_path()
    try:
        with open(file, "r", encoding="utf-8") as f:
            dir = list(yaml.safe_load_all(f))
            path_dir = dir[0]["path"]
            exe_dir = os.path.join(path_dir)
        path = pathlib.Path(exe_dir).parent
        logger.info(f"当前路径:{path}")
        if os.path.exists(os.path.join(path,"winhttp.dll")):
            logger.info("已检测到Modloader，将正常运行")
        else:
            logger.warning("未检测到ModLoader,将自动注入(可能花费几分钟)")
            extract_zip(path)
            time.sleep(2)
            extract_plugin(path)
            
        os.startfile(exe_dir)
        return True
    except FileNotFoundError:
        logger.warning("无法找到游戏路径缓存，将尝试自动定位游戏并建立缓存")
        path = locate_game()
        if path and path != "Not found":
            config = {"path": path}
            with open(get_cache_path(),"w",encoding="utf-8") as f:
                yaml.dump(config, f, default_flow_style=False, allow_unicode=True)
                logger.warning("已建立缓存，程序将自动重启")
                os._exit(0)
        elif path == "Not found":
            logger.error("无法自动查找游戏路径")

    except Exception as e:
        logger.error(f"无法拉起游戏:{e},将自动退出")
        exit(-1)
    return False
logger.info("字幕系统初始化完成")

if pull_up_game():
    logger.info("正在拉起游戏...")
    t = int(w/1920)
    time.sleep(2)
    logger.info("已拉起游戏主进程")
else:
    logger.error("在拉起游戏过程中出现错误，将自动退出。")
    os._exit(-1)
        
def get_window_rect(process_name):
    """
    根据进程名获取窗口坐标
    """
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                if psutil.Process(pid).name() == process_name:
                    hwnds.append(hwnd)
            except psutil.NoSuchProcess:
                pass
        return True
    

    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    
    if hwnds:
        rect = win32gui.GetWindowRect(hwnds[0])
        return rect
    return None

def get_position():
    rect = get_window_rect(ExecutebaleName)
    if rect:
        x,y,w,h = rect
    else:
        return None
    return x,y,w,h


monitoring = False
def monitor():
    global monitoring
    monitoring = True
    logger.info("守护进程已激活")
    while monitoring:
        if check_alive_and_move():
            time.sleep(1)
        else:
            logger.warning("检测到游戏已退出，即将退出程序")
            time.sleep(1)
            os._exit(0)

monitor_thread = threading.Thread(target=monitor)
monitor_thread.daemon = True
monitor_thread.start()

def create_overlay():
    """创建透明悬浮窗"""
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    # 加载UI
    loader = QUiLoader()
    path = pathlib.Path(__file__).resolve().parent
    file = path / "ui.ui"
    window = loader.load(file)
    logger.info("已加载UI文件")
    # 设置透明和置顶
    window.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
    window.setAttribute(Qt.WA_TranslucentBackground)
    window.setAttribute(Qt.WA_TransparentForMouseEvents)
    logger.info("窗口属性设置完成")
    # 简单样式
    window.setStyleSheet("background: transparent; border: none;")
    window.textBrowser.setStyleSheet("""
        background: transparent; 
        color: white; 
        border: none;
        font-size: 14px;
        font-weight: bold;
    """)
    window.textBrowser.setFrameStyle(0)
    
    # 创建信号对象
    signals = OverlaySignals()
    signals.update_text.connect(lambda text: show_text(window, text))
    
    return app, window, signals


def show_text(window, text):
    """显示文本 - 这个函数必须在主线程调用"""
    html = f'''
    <div style="
        color: white; 
        text-shadow: 2px 2px 4px black; 
        padding: 8px;
        background: rgba(0, 0, 0, 0.3);
        border-radius: 5px;
        font-size: 14px;
        font-weight: bold;
    ">{text}</div>
    '''
    window.textBrowser.setHtml(html)
def window_move():
    pos =  get_position()
    if pos:
        px,py,pw,ph = pos
        Pos = QRect(int(px/t),int(py/t),pw,ph)
        window.setGeometry(Pos)
# 初始化悬浮窗
app, window, overlay_signals = create_overlay()
window.show()
logger.info("字幕核心UI初始化完成")
#意义不明のlog
logger.info("程序已启动 (位于5005端口)")
logger.info("正在监听C#插件...")

@server.route("/api", methods=["POST"])
def handle_call():
    try:
        rawdata = json.loads(flask.request.data)
        print(rawdata)
        out = process(raw=rawdata, to_cli=False)
        if out is not None:
            # 使用信号安全地更新UI
            overlay_signals.update_text.emit(out)
        return flask.jsonify({"code": 200, "status": "awa"})
    except Exception as e:
        logger.error(f"服务器出现错误: {e}")
        return flask.jsonify({"code": 404, "status": "man!服务器出问题了"})

if __name__ == "__main__":
    # 在单独的线程中运行Flask服务器
    def run_flask():
        server.run("127.0.0.1", port=5005, debug=False, use_reloader=False)

    flask_thread = threading.Thread(target=run_flask)
    flask_thread.daemon = True
    flask_thread.start()
    
    # 运行Qt主循环
    sys.exit(app.exec())