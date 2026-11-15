# 小睦文件转换器 · Android 版（MuFileTools for Android）

桌面版 [MuFileTools](../README.md) 的 Android 适配版：工具箱风格首页（分类标签 +
工具图标网格，点卡片进入工具子页），界面沿用同一套浅/深色配色与中文文案。

- **应用名**：小睦文件转换器
- **包名**：`com.mutools.filetools`
- **版本**：1.1.3（versionCode 5）
- **系统要求**：Android 8.0（API 26）及以上
- **离线应用**：不申请任何权限、不联网，文件进出全部走系统选择器（SAF / 相册）
- **界面适配**：已处理 Android 15 强制 edge-to-edge（标题栏不顶状态栏、键盘不遮输入框）

---

## 安装

1. 把 `dist_android/MuFileTools-v1.1.3-android.apk` 传到手机（微信/QQ/数据线均可）；
2. 点开安装。系统可能提示「未知来源应用」，允许一次即可；
3. 打开后：首页用分类标签筛选工具，点卡片进入工具子页，左上角「←」或系统返回键回首页；
   右上角按钮循环切换 **跟随系统 → 浅色 → 深色** 外观。

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

### 页面提取（PDF）

| 环节 | 说明 |
| --- | --- |
| 添加 | 「添加 PDF」可多选；顶部下拉在多文件间切换，勾选状态按文件分别记住 |
| 勾选 | 网格缩略图懒加载渲染；点卡片勾选/取消，长按看大图；≤9 页自动全选，否则手动「全选」 |
| 导出格式 | PNG 图片 / JPG 图片 / PNG 长图 / JPG 长图 / PDF 文档 |
| 清晰度 | 100 / 150 / 200 / 300 / 600 DPI（单页超 4000 万像素会自动缩档防崩溃） |
| 文件名 | 支持前缀/后缀；输出目录重名时自动追加 `(1)`、`(2)`（与桌面版 rename 策略一致） |
| 输出位置 | 「选择目录」后记住，之后所有导出直接写入；未选目录时长图/PDF 也可以单文件「另存为」 |

- PNG/JPG：按 DPI 光栅化（PdfRenderer，白底），一页一张图；
- 长图：所选页按最宽页统一宽度竖向拼接，总高超 65000px 时整体缩小（JPEG 单边上限）；
- PDF：用 pdfbox 重新打包选中页，**保留矢量、文字、链接**，不是截图拼的。

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

**Word / PPT / Excel 提取**：桌面版靠 Windows 的 Microsoft Office COM 引擎分页渲染，
Android 没有这个引擎，本版**不做**。已评估三条技术路线，待定方案后加入；
工程按「工具注册制」设计——壳（首页网格 + Host 接口）不感知具体工具，届时新增一个
工具页 + 一条注册即可。

其它差异：无拖放（移动端用文件选择器）、无 DOCX 导出、深浅色记忆在应用内
（跟随系统时随系统自动切换）。

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
        PdfDoc.kt                 PDF 打开/缩略图/按 DPI 渲染/页面提取（pdfbox）
        IcoCore.kt                ICO 位图算法（对应 iconizer.py）
        IcoEncoder.kt             ICO 字节流打包（纯 JVM，可单测）
        ImageOps.kt               EXIF 纠偏解码 + 长图拼接
        Outputs.kt                SAF 写文件、文件名清洗与去重
      ui/
        ExtractPage.kt            页面提取工具页
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

图标资源由 `tools/make_android_icons.py` 从 `app_icon.ico` 生成（mipmap 各密度 +
自适应图标前景，背景色取图标边缘平均色）。
