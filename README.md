# wps2office

双击 `.wps` / `.et` / `.dps` 文件 → 自动按**文件内容**识别真实格式 → 直接用 Word / Excel / PowerPoint 打开**原文件**（不新建副本、不改扩展名、无常驻后台）。

## 新设备安装

1. 把 `wps2office_setup.exe` 拷到新电脑，双击运行一次即可（无需管理员权限）。
2. 安装内容：启动器落位 `%LOCALAPPDATA%\wps2office` → 注册 `.wps/.et/.dps` 文件关联 → 按扩展名显示 Word/Excel/PPT 图标 → 关闭 Excel「格式与扩展名不匹配」警告 → 清理失效的打开方式记录。

## 卸载

```
wps2office_setup.exe --uninstall
```

## 文件说明

| 文件 | 说明 |
|---|---|
| `wps2office_setup.exe` | 单文件安装程序（已内嵌启动器），换设备只需拷这一个文件 |
| `wps2office_setup.py` | 安装程序源码（PyInstaller 打包） |
| `wps2office.py` | 启动器源码（识别逻辑在此） |

## 识别原理

按文件字节特征判断，不信任扩展名：

- OLE2 头（`D0 CF 11 E0`）→ 查内部流名区分 Word（`WordDocument`）/ Excel（`Workbook`）/ PowerPoint（`PowerPoint Document`）
- ZIP 头（`PK`）→ 查 `[Content_Types].xml` 与内部目录区分 docx / xlsx / pptx
- 无法识别 → 弹窗说明，不做任何修改

已知限制：微信聊天窗口内直接点附件会弹微信自带的「打开方式」框（微信自身行为），请先另存到本地再双击。
