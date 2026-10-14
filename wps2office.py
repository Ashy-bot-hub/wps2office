# -*- coding: utf-8 -*-
"""
wps2office — 双击 WPS 系文件（.wps/.et/.dps 及任何改名文件）时，
按【文件内容】自动识别真实格式（Word / Excel / PowerPoint），
直接用对应的 Microsoft Office 程序打开【原文件】。

特点：
  - 不新建任何文件，不修改/重命名原文件
  - 不弹格式选择，双击即开，与正常打开体验一致
  - 无常驻后台：只在双击瞬间运行，开完即退

原理：
  zip 头 PK\\x03\\x04 → 查内部目录 word/ xl/ ppt/ → Word/Excel/PPT 2007+
  OLE2 头 D0CF11E0 → 查 UTF-16LE 流名 WordDocument/Workbook/PowerPoint Document → 97-2003 格式
"""
import sys, os, zipfile

SIG_ZIP = b"PK\x03\x04"
SIG_OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
OLE_STREAMS = [
    (b"W\x00o\x00r\x00d\x00D\x00o\x00c\x00u\x00m\x00e\x00n\x00t\x00", ".doc"),
    (b"W\x00o\x00r\x00k\x00b\x00o\x00o\x00k\x00", ".xls"),
    (b"P\x00o\x00w\x00e\x00r\x00P\x00o\x00i\x00n\x00t\x00 \x00D\x00o\x00c\x00u\x00m\x00e\x00n\x00t\x00", ".ppt"),
]
OFFICE_APPS = {
    ".doc": "WINWORD.EXE", ".docx": "WINWORD.EXE",
    ".xls": "EXCEL.EXE",  ".xlsx": "EXCEL.EXE",
    ".ppt": "POWERPNT.EXE", ".pptx": "POWERPNT.EXE",
}
APP_PATH_KEYS = {
    "WINWORD.EXE": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe",
    "EXCEL.EXE":   r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe",
    "POWERPNT.EXE": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Powerpnt.exe",
}

def detect(path):
    """返回 (格式标签, 扩展名) 或 (标签, None)。只读文件，不做任何修改。"""
    with open(path, "rb") as f:
        head = f.read(8)
        if head[:4] == SIG_ZIP:
            with zipfile.ZipFile(path) as z:
                names = z.namelist()
            if any(n.startswith("word/") for n in names):  return ("Word 2007+ 文档", ".docx")
            if any(n.startswith("xl/")   for n in names):  return ("Excel 2007+ 工作簿", ".xlsx")
            if any(n.startswith("ppt/")  for n in names):  return ("PowerPoint 2007+ 演示文稿", ".pptx")
            return ("zip 压缩包", None)
        if head == SIG_OLE:
            with open(path, "rb") as f:
                blob = f.read()  # 扫全文件，流目录可能在任意扇区
            for pat, ext in OLE_STREAMS:
                if pat in blob:
                    tag = {".doc": "Word 97-2003 文档",
                           ".xls": "Excel 97-2003 工作簿",
                           ".ppt": "PowerPoint 97-2003 演示文稿"}[ext]
                    return (tag, ext)
            return ("OLE2 复合文档（未识别出 Word/Excel/PPT 内容）", None)
    return ("未知格式", None)

def _alert(msg):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, msg, "wps2office", 0x40)
    except Exception:
        print(msg)

def _office_exe(ext):
    """按扩展名定位本机 Office 程序路径（App Paths 注册表 → 常见安装目录）。"""
    exe = OFFICE_APPS[ext]
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, APP_PATH_KEYS[exe]) as k:
            path, _ = winreg.QueryValueEx(k, None)
            if path and os.path.isfile(path):
                return path
    except OSError:
        pass
    for base in (r"C:\Program Files\Microsoft Office\root\Office16",
                 r"C:\Program Files (x86)\Microsoft Office\root\Office16"):
        p = os.path.join(base, exe)
        if os.path.isfile(p):
            return p
    return None

def launch(file_path):
    """双击入口：识别 → 用对应 Office 程序直接打开原文件。"""
    file_path = os.path.abspath(file_path)
    if not os.path.isfile(file_path):
        _alert(f"文件不存在：\n{file_path}"); return 1
    try:
        tag, ext = detect(file_path)
    except Exception as e:
        _alert(f"读取文件失败：{e}"); return 1
    if ext is None:
        _alert(f"无法识别该文件的 Office 格式：\n{os.path.basename(file_path)}\n\n诊断：{tag}\n"
               f"请用 WPS/LibreOffice 打开后另存为标准 Word/Excel/PPT 格式。"); return 1
    exe = _office_exe(ext)
    if exe:
        import subprocess
        subprocess.Popen([exe, file_path])
    else:
        os.startfile(file_path)  # noqa: 找不到 Office 时交给系统默认关联兜底
    return 0

def main():
    args = sys.argv[1:]
    # 启动器模式：exe 运行，或显式 --launch
    if getattr(sys, "frozen", False) or "--launch" in args:
        files = [a for a in args if not a.startswith("-")]
        if not files:
            _alert("请通过双击文件或右键“打开方式”使用本程序。\n\n"
                   "支持：.wps / .et / .dps 及任何被改名的 Office 文件。\n"
                   "程序会按内容自动判断是 Word/Excel/PPT 并直接打开原文件。")
            return 1
        rc = 0
        for p in files:
            rc |= launch(p)
        return rc
    print(__doc__)
    return 1

if __name__ == "__main__":
    sys.exit(main())
