# -*- coding: utf-8 -*-
"""
wps2office_setup — 单文件安装程序。
双击运行即完成：
  1) 释放启动器到 %LOCALAPPDATA%\\wps2office\\wps2office.exe
  2) 注册 .wps / .et / .dps 文件关联（HKCU，无需管理员）
  3) 为各类扩展名配置对应 Office 图标（.wps=Word / .et=Excel / .dps=PowerPoint）
  4) 关闭 Excel「格式与扩展名不匹配」警告（ExtensionHardening=0）
  5) 尽力清除失效的打开方式记录（UserChoice），避免双击弹选择框
之后双击 WPS 系文件即由启动器按内容自动用 Word/Excel/PPT 打开原文件。

卸载：wps2office_setup.exe --uninstall
前提：目标电脑需已安装 Microsoft Office。
"""
import os, sys, shutil, winreg

LAUNCHER_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "wps2office")
LAUNCHER_EXE = os.path.join(LAUNCHER_DIR, "wps2office.exe")
BASE_PROGID = "wps2office.Launcher"
# 扩展名 → (ProgId, 图标来源程序)
EXT_MAP = {
    ".wps": (BASE_PROGID + ".Doc", "WINWORD.EXE"),
    ".et":  (BASE_PROGID + ".Xls", "EXCEL.EXE"),
    ".dps": (BASE_PROGID + ".Ppt", "POWERPNT.EXE"),
}
ALL_PROGIDS = [BASE_PROGID] + [p for p, _ in EXT_MAP.values()]
OFFICE_SEC_KEYS = [
    r"Software\Microsoft\Office\16.0\Excel\Security",
    r"Software\Microsoft\Office\15.0\Excel\Security",
]
APP_PATH_KEYS = {
    "WINWORD.EXE": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe",
    "EXCEL.EXE":   r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe",
    "POWERPNT.EXE": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Powerpnt.exe",
}

def _msg(text, title="wps2office 安装程序"):
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, text, title, 0x40)

def _embedded_launcher():
    if getattr(sys, "frozen", False):
        src = os.path.join(sys._MEIPASS, "wps2office.exe")  # noqa: SLF001
    else:
        src = os.path.join(os.path.dirname(os.path.abspath(__file__)), "wps2office.exe")
    if not os.path.isfile(src):
        raise FileNotFoundError(f"找不到内置启动器：{src}")
    return src

def _setreg(root, path, name, value, rtype=winreg.REG_SZ):
    with winreg.CreateKey(root, path) as k:
        winreg.SetValueEx(k, name, 0, rtype, value)

def _office_icon(exe_name, fallback):
    """取 Office 程序路径作为图标来源（app 图标），找不到则退回启动器图标。"""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, APP_PATH_KEYS[exe_name]) as k:
            path, _ = winreg.QueryValueEx(k, None)
            if path and os.path.isfile(path):
                return f"{path},0"
    except OSError:
        pass
    for base in (r"C:\Program Files\Microsoft Office\root\Office16",
                 r"C:\Program Files (x86)\Microsoft Office\root\Office16"):
        p = os.path.join(base, exe_name)
        if os.path.isfile(p):
            return f"{p},0"
    return fallback

def _refresh_icon_cache():
    """刷新 Windows 图标缓存，让桌面/资源管理器立即显示新图标。"""
    for cmd in ("ie4uinit.exe -show",):
        try:
            os.system(cmd)
        except Exception:
            pass

def install():
    src = _embedded_launcher()
    os.makedirs(LAUNCHER_DIR, exist_ok=True)
    dst = shutil.copy2(src, LAUNCHER_EXE)
    cmd = f'"{dst}" "%1"'
    HKCU = winreg.HKEY_CURRENT_USER

    # 1) 基础 ProgId + Applications 注册（"打开方式"列表可见）
    _setreg(HKCU, rf"Software\Classes\{BASE_PROGID}\shell\open\command", None, cmd)
    _setreg(HKCU, r"Software\Classes\Applications\wps2office.exe\shell\open\command", None, cmd)
    # 2) 按扩展名注册 ProgId + 对应 Office 图标
    for ext, (progid, exe_name) in EXT_MAP.items():
        _setreg(HKCU, rf"Software\Classes\{progid}\shell\open\command", None, cmd)
        _setreg(HKCU, rf"Software\Classes\{progid}\DefaultIcon", None,
                _office_icon(exe_name, f'"{dst},0"'))
        _setreg(HKCU, rf"Software\Classes\{ext}", None, progid)
    # 3) 基础 ProgId 也配图标：老机器上 .et 的 UserChoice 可能记忆指向它
    _setreg(HKCU, rf"Software\Classes\{BASE_PROGID}\DefaultIcon", None,
            _office_icon("EXCEL.EXE", f'"{dst},0"'))
    # 4) 兼容历史 UserChoice 指向场景
    _setreg(HKCU, r"Software\Classes\Applications\EXCEL.EXE\shell\open\command", None, cmd)
    _setreg(HKCU, r"Software\Classes\wps_auto_file\shell\open\command", None, cmd)
    # 5) 关闭 Excel 扩展名不匹配警告
    for sec in OFFICE_SEC_KEYS:
        try:
            _setreg(HKCU, sec, "ExtensionHardening", 0, winreg.REG_DWORD)
        except OSError:
            pass
    # 5) 尽力清除失效/外来的 UserChoice（如残留的 WPS 记录），避免双击弹选择框。
    #    Windows 对 UserChoice 有哈希保护，删除可能被拒绝——被拒则跳过，
    #    用户首次双击时在弹窗中选 wps2office 并勾选“始终”即可（每类扩展名至多一次）。
    for ext in EXT_MAP:
        uc = rf"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{ext}\UserChoice"
        try:
            with winreg.OpenKey(HKCU, uc) as k:
                pid = winreg.QueryValueEx(k, "ProgId")[0]
            if pid == BASE_PROGID or pid in [p for p, _ in EXT_MAP.values()]:
                continue  # 已指向本程序，保留
            winreg.DeleteKey(HKCU, uc)
        except OSError:
            pass
    _refresh_icon_cache()
    _msg(f"安装完成！\n\n启动器位置：{dst}\n已关联：.wps / .et / .dps\n"
         f"双击即按内容自动用 Word / Excel / PPT 打开原文件。\n\n"
         f"提示：若个别文件双击时仍弹出选择框（Windows 哈希保护所致），\n"
         f"请选 wps2office 并勾选“始终使用”，每类扩展名至多操作一次。")

def _del_progid(HKCU, progid):
    for sub in (r"\shell\open\command", r"\shell\open", r"\shell", r"\DefaultIcon", ""):
        try:
            winreg.DeleteKey(HKCU, f"Software\\Classes\\{progid}{sub}")
        except OSError:
            pass

def uninstall():
    HKCU = winreg.HKEY_CURRENT_USER
    for progid in ALL_PROGIDS:
        _del_progid(HKCU, progid)
    _del_progid(HKCU, r"Applications\wps2office.exe")
    for ext, (progid, _) in EXT_MAP.items():
        try:
            with winreg.OpenKey(HKCU, rf"Software\Classes\{ext}") as k:
                if winreg.QueryValueEx(k, None)[0] in ALL_PROGIDS:
                    winreg.DeleteKey(HKCU, rf"Software\Classes\{ext}")
        except OSError:
            pass
    # 用户级覆盖项仅当指向本程序时才移除
    for key in (r"Software\Classes\wps_auto_file\shell\open\command",
                r"Software\Classes\Applications\EXCEL.EXE\shell\open\command"):
        try:
            with winreg.OpenKey(HKCU, key) as k:
                if LAUNCHER_EXE.lower() in str(winreg.QueryValueEx(k, None)[0]).lower():
                    winreg.DeleteValue(k, None)
        except OSError:
            pass
    for base in (r"Software\Classes\wps_auto_file",
                 r"Software\Classes\Applications\EXCEL.EXE"):
        try:
            with winreg.OpenKey(HKCU, base + r"\shell\open\command") as k:
                empty = not winreg.QueryValueEx(k, None)[0]
        except OSError:
            continue
        if empty:
            for sub in (r"\shell\open\command", r"\shell\open", r"\shell", ""):
                try:
                    winreg.DeleteKey(HKCU, base + sub)
                except OSError:
                    pass
    try:
        shutil.rmtree(LAUNCHER_DIR)
    except OSError:
        pass
    _refresh_icon_cache()
    _msg("已卸载：启动器与文件关联已移除。\n原 WPS 文件不会被改动。", "wps2office 卸载")

def main():
    if "--uninstall" in sys.argv[1:]:
        uninstall()
    else:
        install()
    return 0

if __name__ == "__main__":
    sys.exit(main())
