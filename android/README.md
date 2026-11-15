# 小睦文件转换器 · Android 版（MuFileTools for Android）

桌面版 [MuFileTools](../README.md) 的 Android 适配版：工具箱风格首页（分类标签 +
工具图标网格，点卡片进入工具子页），界面沿用同一套浅/深色配色与中文文案。

- **应用名**：小睦文件转换器
- **包名**：`com.mutools.filetools`
- **版本**：1.2.0（versionCode 6）
- **系统要求**：Android 8.0（API 26）及以上
- **离线应用**：不申请任何权限、不联网，文件进出全部走系统选择器（SAF / 相册）
- **界面适配**：已处理 Android 15 强制 edge-to-edge（标题栏不顶状态栏、键盘不遮输入框）

---

## 安装

1. 把 `dist_android/MuFileTools-v1.2.0-android.apk` 传到手机（微信/QQ/数据线均可）；
2. 点开安装。系统可能提示「未知来源应用」，允许一次即可；
3. 打开后：首页用分类标签筛选工具，点卡片进入工具子页，左上角「←」或系统返回键回首页；
   右上角按钮循环切换 **跟随系统 → 浅色 → 深色** 外观。

> **v1.2.0 新增：PPTX 幻灯片 / XLSX 工作表提取。** 不需要装 Office，也不用联网——
> 应用自己拆 `zip` 读 OOXML：PPTX 按「幻灯片 → 版式 → 母版」三级继承拿到形状坐标与字号，
> 当固定版式画到画布上；XLSX 按「共享字符串 + 单元格样式 + 列宽行高」自己画表。
> 于是「页面提取」现在支持 **PDF（按页）/ PPTX（按幻灯片）/ XLSX（按工作表）** 三种输入，
> 导出仍是 PNG / JPG / 长图 / PDF 一套。Word 依然不做，原因见「暂不支持」。

> **v1.1.3 修复**：圆角**从来没真正生效过**（预览里一直显示直角）。
> 根因是 `Bitmap.copy()` 会把源位图的 `hasAlpha` 标志一起复制过去，而相册/相机来的
> JPEG 位图 `hasAlpha=false` —— 往里写 alpha 也没用，绘制时整张按不透明处理。
> 改为新建 ARGB 位图 + `BitmapShader` 填进圆角矩形，圆角外根本不绘制。
> 顺带给大图预览也垫上棋盘格（与桌面版一致）：白底图片的圆角切在浅色卡片上
> 本来也看不出来，棋盘格只在被切掉的四角透出，圆角大小一眼可辨。

> **v1.1.2 修复**：换用新的应用图标（白底手绘女巫），并补上 Android 8.0+ 的
> 自适应图标（`mipmap-anydpi-v26`）—— 此前只给了方形 PNG，各家启动器会自己套白底形状，
> 图标容易被切歪或加一圈白边。现在交给系统按启动器自身形状裁剪，前景与背景色均由
> `tools/make_android_icons.py` 从 `app_icon.ico` 生成。

> **v1.1.1 修复**：图片转 ICO 里**一拖圆角滑杆就闪退**。原因是 `Bitmap.copy(config, false)`
> 得到的是不可变位图，紧接着 `setPixels()` 写入抛 `IllegalStateException`；圆角为 0 时
> 函数会提前返回原图，所以只有"一拖就崩"。同时把圆角蒙版的超采样改为按尺寸自适应
> （原来 1024 图会临时分配 4096×4096 ≈ 64MB 的蒙版），并在换图时回收旧位图，
> 拖动不再累积内存。

> **从 v1.0.0 升级**：v1.0.0 是底部导航切换「页面提取 / 图片转 ICO」，v1.1.0 改成
> 「首页工具网格 + 子页返回栏」——工具数量增加后网格更好扩展（底部导航塞到 5 个就挤了），
> 功能本身没有删减。同时修掉了 v1.0.0 里标题顶进系统状态栏的问题。
> 各版本都用同一个密钥库签名，可直接覆盖安装，不用卸载。

> APK 使用 `android/keystore/mufiletools.keystore` 签名（开发签名）。
> 该密钥库不进仓库，详见「签名」一节；后续版本必须继续用同一个密钥库签名，
> 否则手机上无法覆盖安装。

## 功能

### 首页（工具箱风格）

居中标题 + 分类标签（全部 / 图片处理 / 文档处理）+ 4 列工具图标网格；
点工具卡进入子页（左上角「←」返回，系统返回键同效）。右上角循环切换
**跟随系统 → 浅色 → 深色** 外观。新增工具时在 `MainActivity.tools` 列表登记即可。

### 页面提取（PDF / PPTX / XLSX）

三种输入共用一套界面，区别只在「一个单位是什么」：PDF 是**页**、PPTX 是**幻灯片**、
XLSX 是**工作表**。缩略图角标直接写单位名（`第 3 页` / `销售明细`），计数显示
「已选 3 / 3 个表」，不用记当前是什么格式。

| 环节 | 说明 |
| --- | --- |
| 添加 | 「添加文件」可多选；系统选择器只列 pdf / pptx / xlsx，选到 .ppt / .xls 会提示先另存为新格式 |
| 勾选 | 网格缩略图懒加载渲染；点卡片勾选/取消，长按看大图；≤9 个单位自动全选，否则手动「全选」 |
| 导出格式 | PNG 图片 / JPG 图片 / PNG 长图 / JPG 长图 / PDF 文档 |
| 清晰度 | 100 / 150 / 200 / 300 / 600 DPI（单页超 4000 万像素会自动缩档防崩溃） |
| 文件名 | 支持前缀/后缀；输出目录重名时自动追加 `(1)`、`(2)`（与桌面版 rename 策略一致） |
| 输出位置 | 「选择目录」后记住，之后所有导出直接写入；未选目录时长图/PDF 也可以单文件「另存为」 |

- PNG/JPG：按 DPI 光栅化（白底），一个单位一张图；
- 长图：所选单位按最宽者统一宽度竖向拼接，总高超 65000px 时整体缩小（JPEG 单边上限）；
- PDF：PDF 输入用 pdfbox 重新打包选中页，**保留矢量、文字、链接**；PPTX / XLSX 输入
  没有矢量可复用，是按 DPI 渲染成位图后逐页写入 `PdfDocument`（一页一张整图）。

#### 三种输入怎么解析

| 格式 | 做法 | 已知局限 |
| --- | --- | --- |
| PDF | 系统 `PdfRenderer` 渲染 + pdfbox 取页 | —— |
| PPTX | 读 `presentation.xml` 拿页序与画布尺寸（EMU），逐页把 `p:spTree` 里的形状按 `<a:xfrm>` 画到 Canvas | 只画文本框/占位符、图片、表格、线条、组合与常见预设几何；SmartArt、图表、动画、艺术字不渲染 |
| XLSX | 读 `sheet` + `sharedStrings` + `cellXfs` 样式与 `numFmts`，按列宽行高、合并单元格、边框底纹自己画表 | 不画图表、图片、条件格式图标；公式优先用缓存值，没有缓存值时用内置极简求值器算四则与 SUM/AVERAGE/MIN/MAX/COUNT/ABS/ROUND |

两条硬骨头都记在代码注释里，便于以后维护：

1. **PPT 的坐标和字号常常不在 slide 上**。占位符（`p:ph type/idx`）只写「我是标题」，
   真正的 `off/ext` 和字号在**版式**，甚至版式里也是空 `spPr`，还得再退到**母版**
   （`txStyles` / `titleStyle` / `bodyStyle`）。所以渲染要按「形状 → 版式 → 母版」建继承链，
   母版与版式里的占位符本身**不能画**（它们只是模板）。
2. **EMU 坐标和像素字号不能混在一个坐标系里**。画布保持 EMU（`sldSz` 动辄上亿），
   但文字排版必须先换算成像素——直接在 EMU 量级建 `StaticLayout`，字号会是几十万，
   Skia 在这个量级**一个字节都不出**（画面全白且不报错）。做法是照旧在 EMU 里算位置，
   进文字前 `c.scale(1/s, 1/s)` 抵消外层缩放，之后全按像素画。

PDF 与 PPTX / XLSX 走同一个 `PageSource` 接口（`pageCount` / `label()` / `renderPage()` /
`exportPdf()`），界面层不感知格式；新增一种输入只要实现这个接口并在
`PageSources.open()` 里登记一行。

### 图片转 ICO

与桌面版 core/iconizer.py 逐条对应：

- **选图双入口**：「选择图片…」走系统相册（国产 ROM 的文件选择器常看不到相册），
  「从文件选择」走文件选择器（相册之外/其他应用的图片）；
- 读取失败时提示具体原因（异常类名），解码带 ImageDecoder 后备（AVIF/HEIC 更稳）；
- 读取时纠正 EXIF 方向（手机竖拍不横躺），处理前降采样到 ≤1024px；
- 「居中裁剪」勾选后正方形化，取消则透明像素补边；
- 圆角滑杆 0%（直角）→ 50%（正圆），实时预览，4 倍超采样抗锯齿；
- 预览区 256 / 64 / 32 / 16 棋盘格表示透明区域，是真实像素；
- 导出 `.ico` 含 **256 / 128 / 64 / 48 / 32 / 24 / 16 七档尺寸**：
  256 档按 Windows 规范用 PNG 压缩，其余六档 32 位 BMP，资源管理器直接可用。

### 暂不支持（与桌面版的差异）

**Word 提取**：桌面版靠 Windows 的 Microsoft Office COM 引擎分页，Android 没有这个引擎；
PPT / Excel 之所以能补上，是因为它们**天生就是固定版式或方格**（幻灯片有画布尺寸、
工作表有行列坐标），照坐标画就行，不需要排版引擎。Word 是**流式排版**——一页能装多少字
取决于字体、行高、段距、避头尾、表格跨页规则，等于要在手机上重写一个排版引擎，
所以本版明确不做。要做的话可行路线是引入 LibreOffice 内核，代价是安装包从 3MB 涨到 100MB+。

**旧版二进制格式**（`.ppt` / `.xls` / `.doc`）：OOXML 是 zip + XML，能直接拆；97-2003 的
`.ppt` / `.xls` 是复合文档（OLE）二进制，得另写一套解析。目前选择走系统选择器直接过滤掉，
选到会提示「请先另存为 .pptx / .xlsx」。

其它差异：无拖放（移动端用文件选择器）、无 DOCX / PPTX / XLSX 导出（只导出图片与 PDF）、
深浅色记忆在应用内（跟随系统时随系统自动切换）。

## 构建

本机已验证（Windows + JDK 17 + `D:/Program/Android/SDK`）：

```bash
cd android
# local.properties 里 sdk.dir 指向 Android SDK，按需修改
gradle assembleRelease          # 产物 app/build/outputs/apk/release/app-release.apk
gradle testDebugUnitTest        # ICO 编码器结构自测（纯 JVM，无需设备）
```

本项目未带 gradle wrapper jar，可直接用本机 Gradle 8.9，或在 Android Studio 里
打开 `android/` 目录让它自动配好。

### 签名

仓库里**不含任何密钥库和口令**——公开仓库一旦泄露签名私钥，别人就能伪造你的更新包。
`app/build.gradle.kts` 按下面的顺序读取签名信息，读不到只影响 `assembleRelease`，
`gradle assembleDebug` 始终可用（自动使用 debug 签名）：

1. `android/local.properties`（已在 `.gitignore` 中）：
   ```properties
   keystore.pass=你的密钥库口令
   keystore.alias=mufiletools
   # 可选，默认 ../keystore/mufiletools.keystore
   keystore.path=../keystore/mufiletools.keystore
   ```
2. 环境变量 `MUKEYSTORE_PASS`（CI 或无 local.properties 时）

自己发布请换成自己的密钥库。**同一 App 的后续版本必须用同一密钥库签名**，
否则手机上无法覆盖安装，只能卸载重装（本地偏好会丢）。

## 工程结构

```
android/
  keystore/mufiletools.keystore   签名密钥库（开发用，不进仓库）
  settings.gradle.kts / build.gradle.kts / gradle.properties
  app/
    build.gradle.kts              AGP 8.7.3 · Kotlin 2.0.21 · compileSdk 35 · minSdk 26
    src/main/AndroidManifest.xml  无权限声明
    src/main/java/com/mutools/filetools/
      App.kt                      pdfbox-android 初始化
      MainActivity.kt             壳：工具箱首页（标签+网格）+ 主题 + Host（选择器/建文件路由）
      ui/ToolGridAdapter.kt       工具定义（ToolDef）与首页网格适配器
      core/
        PageSource.kt             统一的「单位来源」接口（页/幻灯片/工作表）+ 按扩展名分发
        Opc.kt                    OOXML 容器层：zip 随机读、.rels 关系解析、XML 小工具
        PdfDoc.kt                 PDF 打开/缩略图/按 DPI 渲染/页面提取（pdfbox）
        PptxDoc.kt                PPTX 解析与渲染（版式/母版继承、形状、图片、表格）
        XlsxDoc.kt                XLSX 解析与渲染（样式、数字/日期格式、合并单元格、公式兜底）
        IcoCore.kt                ICO 位图算法（对应 iconizer.py）
        IcoEncoder.kt             ICO 字节流打包（纯 JVM，可单测）
        ImageOps.kt               EXIF 纠偏解码 + 长图拼接
        Outputs.kt                SAF 写文件、文件名清洗与去重
      ui/
        ExtractPage.kt            页面提取工具页（三种输入共用）
        IcoPage.kt                图片转 ICO 工具页
        PageAdapter.kt            缩略图网格（懒渲染 + LRU 缓存）
      util/Prefs.kt               主题偏好、输出目录记忆
    src/test/                     IcoEncoderTest（ICO 结构字节级校验）
  app/src/main/res/               同桌面版 token 的浅/深色配色、布局、自适应图标
```

## 依赖

| 库 | 用途 |
| --- | --- |
| AndroidX / Material Components | 界面 |
| `android.graphics.pdf.PdfRenderer` | PDF 渲染（系统自带） |
| `com.tom-roush:pdfbox-android` 2.0.27 | PDF 页面提取（Apache-2.0） |
| `androidx.exifinterface` | 照片方向纠正 |
| `androidx.documentfile` | 输出目录写入 |

PPTX / XLSX 的解析**没有引入任何第三方库**：解压用系统 `java.util.zip.ZipFile`、XML 用
`XmlPullParser`、绘制用 `Canvas` / `StaticLayout` / `Typeface`——安装包体积基本没变。

图标资源由 `tools/make_android_icons.py` 从 `app_icon.ico` 生成（mipmap 各密度 +
自适应图标前景，背景色取图标边缘平均色）。
