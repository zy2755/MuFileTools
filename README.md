# 小睦文件转换器 · MuFileTools

一个 Windows 桌面小工具集合，目前集成两个功能：

| 工具 | 功能 |
| --- | --- |
| **页面提取** | 从 **PDF / Word** 里挑出指定的页，从 **Excel** 里挑出指定的工作表，一次导出成**图片**或**文件** |
| **图片转 ICO** | 把 JPG / PNG 等图片变成 Windows `.ico` 图标：圆角实时预览，一键导出七档尺寸 |

界面基于 PySide6，顶部一键切换两个工具。

---

## 快速开始（Windows）

1. **首次使用**：双击 **`install.bat`** —— 自动寻找本机 Python 创建项目自带的 `.venv`，
   并优先用 `wheels\` 里已备好的安装包**离线**装齐依赖（缺包时自动转联网；本仓库不含 wheels 时走网络安装）。
2. **日常启动**：双击 **`run.bat`**（自动用 `.venv` 运行，不依赖系统 PATH 里装了哪个 Python）。
3. 生成演示样例（可选）：`python tools/make_samples.py`

命令行方式：

```powershell
python -m pip install -r requirements.txt   # 或双击 install.bat
python main.py
```

> 为什么需要 `.venv`：本机存在多个 Python 时，直接 `python main.py` 可能把依赖
> 装到了 A 解释器、双击启动却用了 B 解释器（报 `No module named 'PySide6'`）。
> `run.bat` + 项目自带 `.venv` 就是为了避开这个坑。

---

## 工具一：页面提取

### 使用步骤

1. **添加文件**：把 PDF / Word / Excel **拖到窗口任意位置**（拖文件夹会递归扫描子目录），或用「添加文件 / 添加文件夹」。
2. **勾选内容**：左侧选中一个文件，中间会显示这一页/这张表的缩略图；点缩略图或勾选框选择，双击可看大图。页数多的文件默认不勾选，用「全选」。
3. **设置输出**：右侧选格式、清晰度、输出目录，可选文件名前缀/后缀。
4. **开始导出**：进度条走完会提示生成的文件数，可直接打开输出目录。

> 多个文件会依次处理，勾选状态按文件分别记住，一次「开始导出」全部产出。

### 导出格式能做什么

| 输入的选中单位 | 导出格式 | 实际效果 |
| --- | --- | --- |
| PDF / Word 的页 | PNG / JPG | 按 DPI 光栅化，一页一张图；勾「拼成一张长图」则合成竖版长图 |
| PDF / Word 的页 | PDF | 直接复制原始页面对象，矢量、链接、书签都保留 |
| PDF / Word 的页 | DOCX | 每页渲成 100–300 DPI 图片，按 A4 页宽排版，一页一段 |
| Excel 的工作表 | PNG / JPG | 整表渲染（默认「适应一页宽」），宽表自动缩放不被截断 |
| Excel 的工作表 | XLSX | 工作表原样复制到新工作簿，保留格式、列宽、合并单元格、公式 |
| Excel 的工作表 | PDF | 每张表按打印区域转 PDF 后合并 |

---

## 工具二：图片转 ICO

### 使用步骤

1. 点「选择图片…」（或双击左侧预览区，也可直接把图片拖进窗口）；
2. 拖「圆角」滑杆：0% 直角 → 50% 正圆，预览实时刷新；
   非正方形图片可勾选「居中裁剪」，取消则用透明像素补边；
3. 右侧棋盘格代表透明区域，256 / 64 / 32 / 16 四档同时预览；
4. 点「导出 ICO…」保存 `.ico`，内含 **256 / 128 / 64 / 48 / 32 / 24 / 16 七档尺寸**，
   桌面快捷方式、资源管理器、exe 图标替换均可用。

### 设计要点

- **逐档独立缩放**：每个尺寸单独按比例缩放再加圆角（4 倍超采样抗锯齿），小尺寸图标边缘不发糊；
- **透明通道取暗合并**：加圆角不会把图片自身已有的透明区域变成不透明；
- **EXIF 方向纠正**：手机竖拍照片不会横躺；
- **性能**：图像先降采样到 ≤1024px 再处理，拖动滑杆依然流畅；预览自适应窗口大小。

---

## 项目结构

```
main.py                 入口（QApplication + 主窗口壳）
run.bat                 日常启动（自动使用 .venv）
install.bat             首次使用：创建 .venv 并安装依赖
build_exe.bat           PyInstaller 打包成 exe
app_icon.ico            程序图标
core/
  models.py             Unit / ExportSpec / Job 等数据模型 + 支持的文件类型
  sources.py            PdfSource / WordSource / ExcelSource，统一 open/units/thumbnail/render/export
  office.py             Office COM 桥接（Word→PDF、Excel 单表转 PDF/另存）、缓存与降级
  exporter.py           图片 / PDF / DOCX / XLSX 四种导出器 + 长图拼接 + 命名冲突处理
  iconizer.py           ICO 生成核心逻辑（纯 Pillow：读取/正方形化/圆角/打包七档尺寸）
ui/
  shell.py              主窗口壳：品牌栏 + 「页面提取 / 图片转 ICO」切换 + 拖放路由
  dnd.py                全局文件拖放：应用级事件过滤器 + 松手前的悬停提示层
  extract_page.py       页面提取工具页（三栏界面、拖放兜底、导出编排）
  ico_page.py           图片转 ICO 工具页（滑杆实时预览、导出）
  state.py              DocEntry 状态、Loader 与 Thumbnailer 后台线程
  widgets.py            单位卡片、缩略图网格、大图预览
  theme.py              浅色 / 深色两套调色板 + QSS 模板 + 标题栏配色
  settings.py           外观偏好读写（QSettings）
tools/
  make_samples.py       生成 PDF / Word / Excel 演示文件
  make_exe_icon.py      重新生成 app_icon.ico
  fetch_wheels.py       弱网环境分片下载 wheel 再离线安装
tests/
  smoke_test.py         14 项无界面自测（解析→导出→回读校验）
  gui_smoke.py          离屏启动 GUI 的冒烟测试（含 ICO 工具页）
  test_iconizer.py      ICO 核心算法 + 界面冒烟测试
docs/
  设计说明.md            需求拆解、技术选型理由、架构、导出矩阵、已知限制
```

## 外观（深色模式）

右上角 `☾ 深色` / `☀ 浅色` 一键切换，`Ctrl+Shift+D` 同效；「视图」菜单里还可以选
**跟随系统外观 / 浅色 / 深色**。

- **默认跟随 Windows 的应用深浅色**（读注册表 `AppsUseLightTheme`），系统切了它也跟着切；
  一旦你手动点过按钮，就以你的选择为准，不再被系统改回去
- 选择存在 `QSettings`（注册表 `HKCU\Software\MuFileTools`），关掉重开还记得
- 深色不只是换个背景色：**菜单栏、弹出菜单、下拉列表、消息框、标题栏**都跟着变。
  窗口标题栏靠 `DwmSetWindowAttribute` 变成深色，不会留一条白边
- 配色全部集中在 `ui/theme.py` 的 `_LIGHT` / `_DARK` 两个字典里，QSS 里只写
  `{token}` 占位符。想调色改字典即可，不用碰样式模板

## 拖放

**文件拖到窗口任意位置都算数**——包括顶部导航栏、输入框、日志框、下拉框上方，
不用瞄准某个区域。拖文件夹也行，会递归扫描子目录里的同类文件。

拖进来还没松手时，窗口会盖一层提示，直接告诉你松手会发生什么：
「松手即可添加 3 个文件」／「松手即用『图片转 ICO』打开这张图」。

- **跨页自动分流**：在「页面提取」页拖进图片，会自动切到「图片转 ICO」并载入；
  在「图片转 ICO」页拖进文档，会自动切回「页面提取」并加进队列
- 一次只能转一张图，拖多张时载入第一张，其余在状态栏说明忽略了多少
- 不支持的类型会明确说「不支持」，不会静默无视

> 为什么不用各页面自己的 `dragEnterEvent` / `dropEvent`：Qt 的拖放事件只投给
> 鼠标正下方的那个控件，而 `QLineEdit`、`QPlainTextEdit`、`QComboBox`
> **天生就接受拖放**（用来粘文字）。拖到「输出目录」输入框上，文件路径会被
> 当成文字插进框里，永远到不了页面。所以 `ui/dnd.py` 在 `QApplication` 上装了
> 事件过滤器，抢在所有控件之前把「本地文件」的拖放拦下来统一处理。

## 运行环境

- Windows 10/11，Python 3.10+（本机验证 3.13）
- **页面提取建议安装 Microsoft Office**：Word 文件的分页、Excel 工作表的渲染都依赖它
- 没有 Office 时：Excel 仍可用「另存 xlsx」（走 openpyxl，图表/图片不保留）；Word 无法确定页码；装了 LibreOffice 可作 Word→PDF 的备选引擎
- 图片转 ICO 只需 Pillow，无任何系统依赖

## 自测

```powershell
python tools/make_samples.py
python tests/smoke_test.py     # 14 项：三种来源 × 各导出格式，回读校验页数/尺寸/表名
python tests/test_iconizer.py  # ICO 算法 + 界面冒烟
python tests/gui_smoke.py      # 离屏启动界面，检查信号与状态流转（含工具切换）
```

## 打包成 exe（可选）

双击 `build_exe.bat`，用 `.venv` 按 `MuFileTools.spec` 打包，产物在 `dist\MuFileTools.exe`。

两种模式，按用途选：

| 模式 | 命令 | 产物 | 启动耗时 | 适合 |
| --- | --- | --- | --- | --- |
| **单文件** | `build_exe.bat`（即 `MuFileTools.spec`） | `dist\MuFileTools.exe`（约 113 MB） | 约 2.7 s | 发给别人，一个文件搞定 |
| **文件夹** | `PyInstaller MuFileTools_onedir.spec` | `dist_onedir\MuFileTools\`（约 221 MB） | 约 1.3 s | 自己常用，启动快一倍 |

> 单文件模式每次启动都要把内容解压到临时目录，所以启动慢一些，这是 onefile 的固有代价。
> 追求启动速度就用文件夹模式，双击里面的 `MuFileTools.exe` 即可（整个文件夹要一起拷，不能只拿 exe）。

两个 spec 都已剔除运行时用不到的 Qt 组件（软件 OpenGL 兜底 `opengl32sw.dll` 19.7 MB、
Qt 语言包 6.7 MB、TLS 后端、TIFF 格式插件等），详见 spec 内注释。

## 常见问题

**导出的图片糊？** 清晰度选 300 DPI（印刷）或 600 DPI（放大查看）。150 DPI 适合屏幕阅读。

**Word 页码和 Word 里看到的不一致？** 分页由 Word 排版决定。转换前请确认 Word 里没有未更新的域；转换时会调用 `Repaginate` 重排一次。文档改动后缓存自动失效（按修改时间判断）。

**Excel 导出报"未检测到 Microsoft Excel"？** 出图依赖 Excel 打印引擎。可以先装 Office，或改用「Excel 工作簿」格式导出（底层走 openpyxl，不需要 Office）。

**导出很慢？** 首次打开 Word/Excel 需要启动 Office 进程（每文件 1–3 秒）；同一文件再次导出会复用缓存的 PDF。缩略图是懒加载的，滚动时才渲染。

**会动到我正在编辑的 Word/Excel 吗？** 不会。程序用 `DispatchEx` 启动**独立**的 Office 进程，只读打开、关闭时不保存，退出时销毁该进程。

**ICO 导出后小尺寸发糊？** 源图最好 ≥ 256×256。程序对每一档尺寸独立缩放加圆角，已是等比缩放下能做到的最清晰效果；预览区 64/32/16 是真实像素，可提前确认清晰度。

## License

MIT
