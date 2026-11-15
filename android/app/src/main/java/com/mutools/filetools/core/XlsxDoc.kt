package com.mutools.filetools.core

import android.content.Context
import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.Rect
import android.graphics.Typeface
import android.net.Uri
import android.text.Layout
import android.text.SpannableString
import android.text.StaticLayout
import android.text.TextPaint
import android.text.style.AbsoluteSizeSpan
import android.text.style.ForegroundColorSpan
import android.text.style.StyleSpan
import org.xmlpull.v1.XmlPullParser
import java.math.BigDecimal
import java.math.MathContext
import java.text.DecimalFormat
import java.text.SimpleDateFormat
import java.util.Date
import java.util.TimeZone
import kotlin.math.abs
import kotlin.math.floor
import kotlin.math.max
import kotlin.math.min
import kotlin.math.roundToInt

/**
 * XLSX 工作表渲染（纯 Kotlin）。
 *
 * 表格的"一页"不是打印页，而是**工作表（sheet）**——Excel 的分页取决于纸张、缩放、
 * 分页符，本质上不稳定；用户真正想拿走的是"这张表"。所以这里把每个 sheet 画成
 * 一张完整的表格图，长图导出负责纵向拼接。
 *
 * 和 WPS 那套 32×16×7 字节的单元格打包不同，我们不需要为百万行优化：
 * 用稀疏 Map 存单元格，只画用到的区域。样式、列宽、行高、合并、数字格式
 * 都按 OOXML 的原始定义还原，尽量做到"和 Excel 里看到的一样"。
 */
class XlsxDoc(context: Context, uri: Uri) : PageSource {

    // ---------------------------------------------------------------- 模型

    private class SheetRef(val name: String, val part: String) {
        var model: SheetModel? = null
    }

    private class CellData(
        val raw: String,
        val num: Double?,
        val isNumber: Boolean,
        val xf: Int,
        val formula: String? = null,
    )

    private class SheetModel(
        val cells: HashMap<Long, CellData>,
        val rowH: HashMap<Int, Float>,       // 行 → 磅
        val colW: HashMap<Int, Float>,       // 列 → 字符宽
        val merges: List<Rect>,              // 合并区域（0 基，含端点）
        val maxRow: Int,
        val maxCol: Int,
        val defaultRowH: Float,
        val defaultColW: Float,
    )

    private class FontDef(val sz: Float, val bold: Boolean, val italic: Boolean, val color: Int?)

    private class FillDef(val color: Int?)

    private class BorderDef(val left: Int?, val right: Int?, val top: Int?, val bottom: Int?)

    private class XfDef(
        val numFmtId: Int,
        val fontId: Int,
        val fillId: Int,
        val borderId: Int,
        val hAlign: String?,
        val vAlign: String?,
        val wrap: Boolean,
    )

    // ---------------------------------------------------------------- 状态

    private val lock = Any()
    private val opc = Opc.open(context, uri, ".xlsx")
    private val mainPart: String = opc.firstByType("", "officeDocument") ?: "xl/workbook.xml"

    private val sheets = ArrayList<SheetRef>()
    private var sharedStrings: List<String> = emptyList()

    private val fonts = ArrayList<FontDef>()
    private val fills = ArrayList<FillDef>()
    private val borders = ArrayList<BorderDef>()
    private val xfs = ArrayList<XfDef>()
    private val numFmts = HashMap<Int, String>()
    private val themeColors = HashMap<Int, Int>()

    init {
        parseTheme()
        parseSharedStrings()
        parseStyles()
        parseWorkbook()
    }

    override val pageCount: Int get() = sheets.size
    override val unitName: String get() = "个表"
    override fun label(index: Int): String = sheets.getOrNull(index)?.name ?: "表"
    override fun fileLabel(index: Int): String {
        val n = sheets.getOrNull(index)?.name ?: "表${index + 1}"
        return n.replace(Regex("[\\\\/:*?\"<>|]"), "_")
    }

    // ---------------------------------------------------------------- workbook

    private fun parseWorkbook() {
        val p = opc.parser(mainPart) ?: return
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.eventType != XmlPullParser.START_TAG) continue
                if (p.name != "sheet") continue
                val name = p.attr("name") ?: "Sheet${sheets.size + 1}"
                // 隐藏表也列出来：用户可能需要它，只是界面上不显眼
                val part = opc.targetOf(mainPart, p.relId()) ?: continue
                sheets += SheetRef(name, part)
            }
        } catch (_: Exception) {
        }
    }

    private fun parseTheme() {
        val themePart = opc.firstByType(mainPart, "theme") ?: return
        val p = opc.parser(themePart) ?: return
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (!p.at("clrScheme")) continue
                // Excel 的主题索引：0=lt1 1=dk1 2=lt2 3=dk2 4..9=accent 10=hlink 11=folHlink
                val order = listOf("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
                val found = HashMap<String, Int>()
                val depth = p.depth
                while (true) {
                    val ev = p.next()
                    if (ev == XmlPullParser.END_DOCUMENT) break
                    if (ev == XmlPullParser.END_TAG && p.depth == depth) break
                    if (ev != XmlPullParser.START_TAG) continue
                    val slot = p.name
                    val c = readSlotColor(p) ?: continue
                    found[slot] = c
                }
                order.forEachIndexed { i, k -> found[k]?.let { themeColors[i] = it } }
                break
            }
        } catch (_: Exception) {
        }
    }

    /** 主题色槽位（lt1/dk1/accent1…）内部子元素的颜色。 */
    private fun readSlotColor(p: XmlPullParser): Int? {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            if (p.name == "srgbClr" || p.name == "sysClr") return parseThemeColorEl(p)
        }
        return null
    }

    private fun parseThemeColorEl(p: XmlPullParser): Int? {
        val v = when (p.name) {
            "srgbClr" -> p.attr("val")?.let { parseColor(it, Color.BLACK) }
            "sysClr" -> (p.attr("lastClr") ?: p.attr("val"))?.let { parseColor(it, Color.BLACK) }
            else -> null
        }
        p.skip()
        return v
    }

    private fun parseSharedStrings() {
        val part = opc.firstByType(mainPart, "sharedStrings") ?: return
        val p = opc.parser(part) ?: return
        val list = ArrayList<String>()
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (!p.at("si")) continue
                list += readRichText(p)
            }
        } catch (_: Exception) {
        }
        sharedStrings = list
    }

    /** `<si>` 或 `<is>`：可能是单段 `<t>`，也可能是多段富文本 `<r><t>`。 */
    private fun readRichText(p: XmlPullParser): String {
        val sb = StringBuilder()
        val depth = p.depth
        var phoneticDepth = -1
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "t" -> if (phoneticDepth < 0) sb.append(p.plainText())
                // rPh 是注音，混进正文会变成乱码
                "rPh" -> { phoneticDepth = p.depth; p.skip(); phoneticDepth = -1 }
            }
        }
        return sb.toString()
    }

    // ---------------------------------------------------------------- styles

    private fun parseStyles() {
        val part = opc.firstByType(mainPart, "styles") ?: return
        val p = opc.parser(part) ?: return
        // cellStyleXfs 里也有 <xf>，但单元格的 s 索引指的是 cellXfs，两者不能混
        var inCellXfs = false
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.eventType == XmlPullParser.END_TAG) {
                    if (p.name == "cellXfs") inCellXfs = false
                    continue
                }
                when {
                    p.at("cellXfs") -> inCellXfs = true
                    p.at("numFmt") -> {
                        val id = p.attrInt("numFmtId", -1)
                        val code = p.attr("formatCode")
                        if (id >= 0 && code != null) numFmts[id] = code
                    }
                    p.at("font") -> fonts += parseFont(p)
                    p.at("fill") -> fills += parseFill(p)
                    p.at("border") -> borders += parseBorder(p)
                    p.at("xf") -> if (inCellXfs) xfs += parseXf(p) else p.skip()
                }
            }
        } catch (_: Exception) {
        }
    }

    private fun parseFont(p: XmlPullParser): FontDef {
        var sz = 11f
        var bold = false
        var italic = false
        var color: Int? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "sz" -> sz = p.attrFloat("val", 11f)
                "b" -> bold = p.attr("val") != "0" && p.attr("val") != "false"
                "i" -> italic = p.attr("val") != "0" && p.attr("val") != "false"
                "color" -> color = colorOf(p)
                else -> {}
            }
        }
        return FontDef(sz, bold, italic, color)
    }

    private fun colorOf(p: XmlPullParser): Int? {
        val rgb = p.attr("rgb")
        if (rgb != null) return parseColor(rgb, Color.BLACK)
        val theme = p.attr("theme")?.toIntOrNull()
        if (theme != null) {
            themeColors[theme]?.let { return it }
            return Color.BLACK
        }
        val indexed = p.attr("indexed")?.toIntOrNull()
        if (indexed != null) return indexedColor(indexed)
        return null
    }

    /** 调色板里 0/1 黑白、2 红、3 绿、4 蓝…（老文件会用 indexed） */
    private fun indexedColor(i: Int): Int = when (i) {
        0, 1 -> Color.BLACK
        2 -> Color.WHITE
        3 -> 0xFFFF0000.toInt()
        4 -> 0xFF00FF00.toInt()
        5 -> 0xFF0000FF.toInt()
        6 -> 0xFFFFFF00.toInt()
        7 -> 0xFFFF00FF.toInt()
        8 -> 0xFF00FFFF.toInt()
        else -> Color.BLACK
    }

    private fun parseFill(p: XmlPullParser): FillDef {
        var color: Int? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "patternFill" -> if (p.attr("patternType") == "solid") {
                    val d2 = p.depth
                    while (true) {
                        val e2 = p.next()
                        if (e2 == XmlPullParser.END_DOCUMENT) break
                        if (e2 == XmlPullParser.END_TAG && p.depth == d2) break
                        if (e2 != XmlPullParser.START_TAG) continue
                        if (p.name == "fgColor") color = colorOf(p)
                    }
                }
                "gradientFill" -> color = 0xFFEFF3F8.toInt()   // 渐变退化为浅底
            }
        }
        return FillDef(color)
    }

    private fun parseBorder(p: XmlPullParser): BorderDef {
        var l: Int? = null; var r: Int? = null; var t: Int? = null; var b: Int? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            val style = p.attr("style")
            val hasBorder = style != null && style != "none"
            if (!hasBorder) {
                if (p.name in arrayOf("left", "right", "top", "bottom")) p.skip()
                continue
            }
            val color = pickBorderColor(p)
            when (p.name) {
                "left" -> l = color
                "right" -> r = color
                "top" -> t = color
                "bottom" -> b = color
            }
        }
        return BorderDef(l, r, t, b)
    }

    private fun pickBorderColor(p: XmlPullParser): Int {
        var color = 0xFFBFBFBF.toInt()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            if (p.name == "color") color = colorOf(p) ?: color
        }
        return color
    }

    private fun parseXf(p: XmlPullParser): XfDef {
        val numFmtId = p.attrInt("numFmtId", 0)
        val fontId = p.attrInt("fontId", 0)
        val fillId = p.attrInt("fillId", 0)
        val borderId = p.attrInt("borderId", 0)
        var hAlign: String? = null
        var vAlign: String? = null
        var wrap = false
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            if (p.name == "alignment") {
                hAlign = p.attr("horizontal")
                vAlign = p.attr("vertical")
                wrap = p.attr("wrapText") == "1" || p.attr("wrapText") == "true"
            }
        }
        return XfDef(numFmtId, fontId, fillId, borderId, hAlign, vAlign, wrap)
    }

    // ---------------------------------------------------------------- worksheet

    private fun modelOf(index: Int): SheetModel = synchronized(lock) {
        val ref = sheets[index]
        ref.model?.let { return it }
        val m = parseSheet(ref.part)
        ref.model = m
        m
    }

    private fun parseSheet(part: String): SheetModel {
        val cells = HashMap<Long, CellData>()
        val rowH = HashMap<Int, Float>()
        val colW = HashMap<Int, Float>()
        val merges = ArrayList<Rect>()
        var maxRow = 0
        var maxCol = 0
        var defaultRowH = 15f
        var defaultColW = 8.43f
        var curRow = -1

        val p = opc.parser(part) ?: return SheetModel(cells, rowH, colW, merges, 0, 0, 15f, 8.43f)
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                when {
                    p.at("sheetFormatPr") -> {
                        defaultRowH = p.attrFloat("defaultRowHeight", defaultRowH)
                        defaultColW = p.attrFloat("baseColWidth", defaultColW)
                    }
                    p.at("col") -> {
                        val minC = p.attrInt("min", 1) - 1
                        val maxC = p.attrInt("max", minC + 1) - 1
                        if (p.attr("hidden") == "1") {
                            for (k in minC..min(maxC, minC + 255)) colW[k] = 0f
                        } else {
                            val w = p.attrFloat("width", defaultColW)
                            for (k in minC..min(maxC, minC + 255)) colW[k] = w
                        }
                    }
                    p.at("row") -> {
                        curRow = p.attrInt("r", curRow + 2) - 1
                        val ht = p.attr("ht")?.toFloatOrNull()
                        if (ht != null && ht > 0f) rowH[curRow] = ht
                        maxRow = max(maxRow, curRow)
                    }
                    p.at("c") -> {
                        val ref = p.attr("r") ?: ""
                        val (r, c) = refToCell(ref) ?: (curRow to -1)
                        val col = if (c >= 0) c else 0
                        val row = if (r >= 0) r else curRow
                        if (row >= 0) {
                            val cell = parseCell(p)
                            if (cell != null) {
                                cells[key(row, col)] = cell
                                maxRow = max(maxRow, row)
                                maxCol = max(maxCol, col)
                            }
                        }
                    }
                    p.at("mergeCell") -> {
                        val ref = p.attr("ref")
                        if (ref != null) rangeOf(ref)?.let {
                            merges += it
                            maxRow = max(maxRow, it.bottom)
                            maxCol = max(maxCol, it.right)
                        }
                    }
                }
            }
        } catch (_: Exception) {
        }
        return SheetModel(cells, rowH, colW, merges, maxRow, maxCol, defaultRowH, defaultColW)
    }

    private fun parseCell(p: XmlPullParser): CellData? {
        val t = p.attr("t") ?: "n"
        val s = p.attrInt("s", 0)
        var v = ""
        var inline = ""
        var formula: String? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "v" -> v = p.plainText()
                "is" -> inline = readRichText(p)
                "f" -> formula = p.plainText()
            }
        }
        return when (t) {
            "s" -> {
                val idx = v.trim().toIntOrNull() ?: return null
                CellData(sharedStrings.getOrElse(idx) { "" }, null, false, s)
            }
            "inlineStr" -> if (inline.isEmpty()) null else CellData(inline, null, false, s)
            "str" -> if (v.isEmpty()) null else CellData(v, null, false, s)
            "b" -> CellData(if (v.trim() == "1") "TRUE" else "FALSE", null, false, s)
            "e" -> CellData(v, null, false, s)
            "d" -> CellData(v, null, false, s)
            else -> {
                val d = v.trim().toDoubleOrNull()
                // 只有公式、没有缓存值（openpyxl 等生成的文件很常见）——先留着，渲染时再算
                if (d == null && formula != null) CellData("", null, true, s, formula)
                else if (d == null) null
                else CellData(v, d, true, s, formula)
            }
        }
    }

    // ---------------------------------------------------------------- 渲染

    override fun renderPage(index: Int, targetWidth: Int): Bitmap = synchronized(lock) {
        val model = modelOf(index)
        val geo = geometry(model)
        val scale = (targetWidth.coerceAtLeast(32) / geo.width).coerceAtLeast(0.05f)
        drawSheet(index, model, geo, scale)
    }

    override fun renderPageDpi(index: Int, dpi: Int): Bitmap = synchronized(lock) {
        val model = modelOf(index)
        val geo = geometry(model)
        var scale = dpi / 96f
        val area = geo.width * scale * geo.height * scale
        if (area > MAX_PIXELS) scale *= kotlin.math.sqrt(MAX_PIXELS / area)
        drawSheet(index, model, geo, scale)
    }

    /** 版面几何，单位是 96 DPI 的像素。 */
    private class Geo(
        val colX: FloatArray, val colW: FloatArray,
        val rowY: FloatArray, val rowH: FloatArray,
        val width: Float, val height: Float,
    )

    private fun geometry(m: SheetModel): Geo {
        val cols = m.maxCol + 1
        val rows = m.maxRow + 1
        val colW = FloatArray(cols)
        val rowH = FloatArray(rows)

        // 列宽：有显式宽度就照抄；没有就看内容最宽的一行，避免文字被切掉
        for (c in 0 until cols) {
            val w = m.colW[c]
            colW[c] = when {
                w != null && w <= 0f -> 0f                        // 隐藏列
                w != null -> w * CHAR_PX + COL_PAD
                else -> autoColWidth(m, c) ?: (m.defaultColW * CHAR_PX + COL_PAD)
            }
            colW[c] = colW[c].coerceIn(6f, 420f)
        }
        for (r in 0 until rows) {
            val h = m.rowH[r]
            rowH[r] = if (h != null && h > 0f) h * 96f / 72f else m.defaultRowH * 96f / 72f
            rowH[r] = rowH[r].coerceIn(8f, 800f)
        }

        val colX = FloatArray(cols + 1)
        for (c in 0 until cols) colX[c + 1] = colX[c] + colW[c]
        val rowY = FloatArray(rows + 1)
        for (r in 0 until rows) rowY[r + 1] = rowY[r] + rowH[r]
        return Geo(colX, colW, rowY, rowH, max(1f, colX[cols]), max(1f, rowY[rows]))
    }

    private fun autoColWidth(m: SheetModel, col: Int): Float? {
        var widest = 0f
        var found = false
        for ((k, cell) in m.cells) {
            if ((k % KEY_STRIDE).toInt() != col) continue
            found = true
            val text = displayText(cell, m)
            if (text.isEmpty()) continue
            val w = estimateWidth(text)
            if (w > widest) widest = w
            if (widest > 400f) break
        }
        if (!found) return null
        return (widest + COL_PAD).coerceIn(20f, 420f)
    }

    /**
     * 粗略估宽：中文按 1 个全角宽、西文按约 0.55 个字符宽。
     * 这里宁可估宽一点——估窄了文字会被省略号截断，那是没法看的。
     */
    private fun estimateWidth(text: String): Float {
        var w = 0f
        for (ch in text.take(200)) {
            w += if (ch.code > 0x2E80) 17f else 8.5f
        }
        return w
    }

    private fun displayText(cell: CellData, m: SheetModel?): String {
        if (!cell.isNumber) return cell.raw
        // 缓存值优先；没有就现场算一遍公式（Excel 会重算，我们只做四则与常见聚合）
        val value = cell.num
            ?: cell.formula?.let { f -> m?.let { evalFormula(it, f, 0) } }
            ?: return ""
        val xf = xfs.getOrNull(cell.xf)
        val code = xf?.let { numFmts[it.numFmtId] ?: builtinFormat(it.numFmtId) } ?: "General"
        return formatNumber(value, code)
    }

    /** drawCell 里没有 model 参数，这里存一份"当前正在画的工作表"，避免到处传参。 */
    private var drawingModel: SheetModel? = null

    private fun drawSheet(index: Int, m: SheetModel, geo: Geo, scale: Float): Bitmap {
        drawingModel = m
        val w = max(16, (geo.width * scale).roundToInt())
        val h = max(16, (geo.height * scale).roundToInt())
        val bmp = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888)
        val c = Canvas(bmp)
        c.drawColor(Color.WHITE)
        c.scale(scale, scale)   // 之后一律用 96DPI 像素坐标

        // 合并区域：先落一份"被覆盖"集合，避免重复画内容
        val covered = HashSet<Long>()
        val mergeByOrigin = HashMap<Long, Rect>()
        for (r in m.merges) {
            mergeByOrigin[key(r.top, r.left)] = r
            for (rr in r.top..r.bottom) for (cc in r.left..r.right) {
                if (rr != r.top || cc != r.left) covered += key(rr, cc)
            }
        }

        val borderPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply { style = Paint.Style.STROKE }

        // 先铺一层极淡的网格线（Excel 的空白格线），有内容的单元格会盖在上面
        val grid = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            color = 0xFFE9ECF1.toInt()
            style = Paint.Style.STROKE
            strokeWidth = 1f
        }
        for (ri in 0 until geo.rowY.size - 1) {
            c.drawLine(0f, geo.rowY[ri], geo.width, geo.rowY[ri], grid)
        }
        for (ci in 0 until geo.colX.size - 1) {
            if (geo.colW[ci] <= 0.5f) continue
            c.drawLine(geo.colX[ci], 0f, geo.colX[ci], geo.height, grid)
        }

        for ((k, cell) in m.cells) {
            if (k in covered) continue
            val row = (k / KEY_STRIDE).toInt()
            val col = (k % KEY_STRIDE).toInt()
            if (row < 0 || row >= geo.rowY.size - 1 || col < 0 || col >= geo.colX.size - 1) continue

            val merge = mergeByOrigin[k]
            val right = if (merge != null) merge.right.coerceAtMost(geo.colX.size - 2) else col
            val bottom = if (merge != null) merge.bottom.coerceAtMost(geo.rowY.size - 2) else row
            val rect = Rect(
                geo.colX[col].roundToInt(), geo.rowY[row].roundToInt(),
                geo.colX[right + 1].roundToInt(), geo.rowY[bottom + 1].roundToInt(),
            )
            if (rect.width() <= 0 || rect.height() <= 0) continue

            val xf = xfs.getOrNull(cell.xf)
            drawCell(c, rect, cell, xf, borderPaint)
        }
        return bmp
    }

    private fun drawCell(c: Canvas, rect: Rect, cell: CellData, xf: XfDef?, borderPaint: Paint) {
        val font = fonts.getOrNull(xf?.fontId ?: 0) ?: FontDef(11f, false, false, null)
        val fill = fills.getOrNull(xf?.fillId ?: 0)?.color
        if (fill != null && (fill ushr 24) != 0) {
            c.drawRect(
                rect.left.toFloat(), rect.top.toFloat(), rect.right.toFloat(), rect.bottom.toFloat(),
                Paint().apply { color = fill }
            )
        }

        // 边框：只画有定义的那几条边，没定义就留白（Excel 的"无边框"）
        val bd = borders.getOrNull(xf?.borderId ?: 0)
        if (bd != null) {
            borderPaint.strokeWidth = 1.2f
            val l = rect.left.toFloat(); val t = rect.top.toFloat()
            val r = rect.right.toFloat(); val b = rect.bottom.toFloat()
            bd.left?.let { borderPaint.color = it; c.drawLine(l, t, l, b, borderPaint) }
            bd.right?.let { borderPaint.color = it; c.drawLine(r, t, r, b, borderPaint) }
            bd.top?.let { borderPaint.color = it; c.drawLine(l, t, r, t, borderPaint) }
            bd.bottom?.let { borderPaint.color = it; c.drawLine(l, b, r, b, borderPaint) }
        }

        val text = displayText(cell, drawingModel)
        if (text.isEmpty()) return

        val sz = font.sz * 96f / 72f
        val color = font.color ?: Color.BLACK
        val pad = 3f
        val availW = rect.width() - pad * 2
        if (availW <= 0) return

        val align = when (xf?.hAlign) {
            "center", "centerContinuous" -> Layout.Alignment.ALIGN_CENTER
            "right" -> Layout.Alignment.ALIGN_OPPOSITE
            "left" -> Layout.Alignment.ALIGN_NORMAL
            else -> if (cell.isNumber) Layout.Alignment.ALIGN_OPPOSITE else Layout.Alignment.ALIGN_NORMAL
        }

        val sp = SpannableString(text)
        sp.setSpan(AbsoluteSizeSpan(sz.roundToInt().coerceAtLeast(1)), 0, text.length, 0)
        sp.setSpan(ForegroundColorSpan(color), 0, text.length, 0)
        if (font.bold || font.italic) {
            sp.setSpan(
                StyleSpan(
                    when {
                        font.bold && font.italic -> Typeface.BOLD_ITALIC
                        font.bold -> Typeface.BOLD
                        else -> Typeface.ITALIC
                    }
                ), 0, text.length, 0
            )
        }
        val paint = TextPaint(Paint.ANTI_ALIAS_FLAG).apply { textSize = sz }

        val wrap = xf?.wrap == true
        val layout = StaticLayout.Builder
            .obtain(sp, 0, sp.length, paint, availW.roundToInt().coerceAtLeast(1))
            .setAlignment(align)
            .setIncludePad(false)
            .setMaxLines(if (wrap) 40 else 1)
            .setEllipsize(if (wrap) null else android.text.TextUtils.TruncateAt.END)
            .build()

        val ty = when (xf?.vAlign) {
            "top" -> rect.top + pad
            "center" -> rect.top + (rect.height() - layout.height) / 2f
            else -> rect.bottom - layout.height - pad   // Excel 默认底部对齐
        }.coerceAtLeast(rect.top.toFloat())
        c.save()
        c.translate(rect.left + pad, ty)
        c.clipRect(0, 0, availW.roundToInt(), (rect.bottom - ty).coerceAtLeast(1f).roundToInt())
        layout.draw(c)
        c.restore()
    }

    override fun close() {
        opc.close()
    }

    // ---------------------------------------------------------------- 数字格式

    private fun builtinFormat(id: Int): String = when (id) {
        0 -> "General"
        1 -> "0"
        2 -> "0.00"
        3 -> "#,##0"
        4 -> "#,##0.00"
        9 -> "0%"
        10 -> "0.00%"
        11 -> "0.00E+00"
        14 -> "m/d/yy"
        15 -> "d-mmm-yy"
        16 -> "d-mmm"
        17 -> "mmm-yy"
        18 -> "h:mm AM/PM"
        19 -> "h:mm:ss AM/PM"
        20 -> "h:mm"
        21 -> "h:mm:ss"
        22 -> "m/d/yy h:mm"
        37, 38, 39, 40 -> "#,##0"
        45 -> "mm:ss"
        46 -> "[h]:mm:ss"
        47 -> "mmss.0"
        49 -> "@"
        else -> "General"
    }

    private fun formatNumber(v: Double, code: String): String {
        if (code == "General" || code.isEmpty()) return general(v)
        // 多段格式用 ';' 分开：正数;负数;零;文本
        val sections = splitSections(code)
        val section = when {
            v > 0 && sections.isNotEmpty() -> sections[0]
            v < 0 && sections.size > 1 -> sections[1]
            v == 0.0 && sections.size > 2 -> sections[2]
            else -> sections.getOrElse(0) { "General" }
        }
        if (section.isEmpty()) return ""
        if (isDateFormat(section)) return formatDate(v, section)

        val percent = section.contains('%')
        val scale = if (percent) 100.0 else 1.0
        val abs = abs(v * scale)

        val dot = section.lastIndexOf('.')
        val decimals = if (dot >= 0) {
            section.substring(dot + 1).count { it == '0' || it == '#' }
        } else 0
        val grouping = Regex("[0#?],[0#?]").containsMatchIn(section)

        val pattern = buildString {
            if (grouping) append("#,##0") else append("0")
            if (decimals > 0) append(".").append("0".repeat(decimals.coerceAtMost(10)))
        }
        val formatted = try {
            DecimalFormat(pattern).format(abs)
        } catch (_: Exception) {
            general(abs)
        }

        val (prefix, suffix) = literals(section)
        val sign = if (v < 0 && sections.size <= 1) "-" else ""
        return sign + prefix + formatted + suffix
    }

    private fun splitSections(code: String): List<String> {
        val out = ArrayList<String>()
        val sb = StringBuilder()
        var inQuote = false
        for (ch in code) {
            when {
                ch == '"' -> { inQuote = !inQuote; sb.append(ch) }
                ch == ';' && !inQuote -> { out += sb.toString(); sb.setLength(0) }
                else -> sb.append(ch)
            }
        }
        out += sb.toString()
        return out
    }

    /** 取出格式里的字面量（货币符号、"工作日"这类），返回前缀/后缀。 */
    private fun literals(code: String): Pair<String, String> {
        val clean = StringBuilder()
        val isPh = ArrayList<Boolean>()
        fun add(s: String, ph: Boolean) {
            for (c in s) { clean.append(c); isPh.add(ph) }
        }
        var i = 0
        while (i < code.length) {
            val ch = code[i]
            when {
                ch == '"' -> {
                    val end = code.indexOf('"', i + 1).let { if (it < 0) code.length else it }
                    add(code.substring(i + 1, end), false)
                    i = end + 1
                }
                ch == '\\' && i + 1 < code.length -> { add(code[i + 1].toString(), false); i += 2 }
                ch == '_' || ch == '*' -> i += 2                     // 占位空格/填充字符
                ch == '[' -> {                                       // [Red] [$-804] [>=100]
                    val end = code.indexOf(']', i).let { if (it < 0) code.length - 1 else it }
                    val inner = code.substring(i + 1, end)
                    // 货币区域 [$¥-804] 里把符号取出来
                    Regex("""^\$(.)""").find(inner)?.groupValues?.getOrNull(1)?.let { add(it, false) }
                    i = end + 1
                }
                ch in "0#?.," -> { add(ch.toString(), true); i++ }
                else -> { add(ch.toString(), false); i++ }
            }
        }
        val first = isPh.indexOfFirst { it }
        if (first < 0) return clean.toString() to ""
        val last = isPh.indexOfLast { it }
        return clean.substring(0, first) to clean.substring(last + 1)
    }

    private fun general(v: Double): String {
        if (v == floor(v) && abs(v) < 1e11) return v.toLong().toString()
        val bd = BigDecimal(v).round(MathContext(10)).stripTrailingZeros()
        return bd.toPlainString()
    }

    private fun isDateFormat(code: String): Boolean {
        // 去掉引号里的字面量、方括号（颜色/条件）再判断
        var s = code.replace(Regex("\"[^\"]*\""), "").replace(Regex("\\[[^\\]]*\\]"), "")
        s = s.replace("General", "")
        return Regex("[yYmMdDhHsS]").containsMatchIn(s)
    }

    private fun formatDate(serial: Double, code: String): String {
        // Excel 序列日 → 1970 基准：25569 = 1900-01-01 到 1970-01-01 的天数
        val ms = ((serial - 25569.0) * 86400000.0).toLong()
        val javaPattern = toJavaPattern(code)
        return try {
            SimpleDateFormat(javaPattern, java.util.Locale.getDefault()).apply {
                timeZone = TimeZone.getTimeZone("UTC")
            }.format(Date(ms))
        } catch (_: Exception) {
            general(serial)
        }
    }

    /** Excel 的 m 既可能是月也可能是分：紧挨 h/s 时按分钟处理。 */
    private fun toJavaPattern(code: String): String {
        val sb = StringBuilder()
        var i = 0
        var lastWasHour = false
        while (i < code.length) {
            val ch = code[i]
            var n = 1
            while (i + n < code.length && code[i + n] == ch) n++
            val lower = ch.lowercaseChar()
            when {
                lower == 'y' -> sb.append("y".repeat(n))
                lower == 'd' -> sb.append("d".repeat(n))
                lower == 'h' -> { sb.append("H".repeat(n)); lastWasHour = true }
                lower == 's' -> { sb.append("s".repeat(n)); lastWasHour = false }
                lower == 'm' -> {
                    val nextIsSec = code.getOrNull(i + n)?.lowercaseChar() == 's'
                    val asMinute = lastWasHour || nextIsSec
                    sb.append((if (asMinute) "m" else "M").repeat(n))
                    lastWasHour = false
                }
                code.startsWith("AM/PM", i, true) || code.startsWith("A/P", i, true) -> {
                    sb.append("a")
                    i += if (code.startsWith("AM/PM", i, true)) 5 else 3
                    continue
                }
                ch == '[' -> {
                    val end = code.indexOf(']', i).let { if (it < 0) code.length - 1 else it }
                    val inner = code.substring(i + 1, end)
                    if (inner.lowercase() == "h") sb.append("H")
                    else if (inner.lowercase() == "m") sb.append("m")
                    else if (inner.lowercase() == "s") sb.append("s")
                    i = end + 1
                    continue
                }
                ch == '"' -> {
                    val end = code.indexOf('"', i + 1).let { if (it < 0) code.length else it }
                    sb.append("'").append(code, i + 1, end).append("'")
                    i = end + 1
                    continue
                }
                ch == '\\' && i + 1 < code.length -> { sb.append("'").append(code[i + 1]).append("'"); i += 2; continue }
                ch in ".\\/-: ," -> sb.append("'").append(ch).append("'")
                else -> {}
            }
            i += n
        }
        return sb.toString().ifEmpty { "yyyy-MM-dd" }
    }

    // ---------------------------------------------------------------- 公式求值

    /**
     * 极简求值：四则、括号、乘方、单元格引用，外加 SUM/AVERAGE/MIN/MAX/COUNT/ABS/ROUND。
     * 只用于"文件里没存缓存值"的情形（openpyxl 之类生成的文件很常见），
     * 算不出来就留空，绝不瞎猜。
     */
    private fun evalFormula(m: SheetModel, formula: String, depth: Int): Double? {
        if (depth > 6) return null
        val src = formula.trim().removePrefix("=")
        if (src.isEmpty()) return null
        return Expr(m, src, depth).parseAll()
    }

    private inner class Expr(private val m: SheetModel, private val src: String, private val depth: Int) {
        private var i = 0

        fun parseAll(): Double? {
            val v = sum()
            skipWs()
            return if (i >= src.length) v else null
        }

        private fun skipWs() {
            while (i < src.length && src[i].isWhitespace()) i++
        }

        private fun peek(): Char? {
            skipWs()
            return if (i < src.length) src[i] else null
        }

        private fun sum(): Double? {
            var v = product() ?: return null
            while (true) {
                when (peek()) {
                    '+' -> { i++; v += product() ?: return null }
                    '-' -> { i++; v -= product() ?: return null }
                    else -> return v
                }
            }
        }

        private fun product(): Double? {
            var v = power() ?: return null
            while (true) {
                when (peek()) {
                    '*' -> { i++; v *= power() ?: return null }
                    '/' -> {
                        i++
                        val d = power() ?: return null
                        if (d == 0.0) return null
                        v /= d
                    }
                    else -> return v
                }
            }
        }

        private fun power(): Double? {
            val base = unary() ?: return null
            if (peek() == '^') {
                i++
                val e = power() ?: return null
                return Math.pow(base, e)
            }
            return base
        }

        private fun unary(): Double? = when (peek()) {
            '-' -> { i++; val v = unary() ?: return null; -v }
            '+' -> { i++; unary() }
            '(' -> { i++; val v = sum() ?: return null; if (peek() == ')') i++; v }
            else -> atom()
        }

        private fun atom(): Double? {
            skipWs()
            if (i >= src.length) return null
            val c = src[i]
            return when {
                c.isDigit() || c == '.' -> number()
                c.isLetter() || c == '$' -> nameOrRef()
                else -> null
            }
        }

        private fun number(): Double? {
            val start = i
            while (i < src.length && (src[i].isDigit() || src[i] == '.')) i++
            return src.substring(start, i).toDoubleOrNull()
        }

        private fun nameOrRef(): Double? {
            val start = i
            while (i < src.length && (src[i].isLetterOrDigit() || src[i] == '$' || src[i] == '_')) i++
            val name = src.substring(start, i).uppercase()
            if (peek() == '(') return call(name)
            val cell = refToCell(name.replace("$", "")) ?: return null
            return cellValue(cell.first, cell.second)
        }

        private fun cellValue(row: Int, col: Int): Double? {
            val c = m.cells[key(row, col)] ?: return null
            c.num?.let { return it }
            val f = c.formula ?: return null
            return evalFormula(m, f, depth + 1)
        }

        private fun call(name: String): Double? {
            i++ // '('
            val args = ArrayList<String>()
            val sb = StringBuilder()
            var paren = 0
            while (i < src.length) {
                val ch = src[i]
                when {
                    ch == '(' -> { paren++; sb.append(ch); i++ }
                    ch == ')' && paren > 0 -> { paren--; sb.append(ch); i++ }
                    ch == ')' -> { args += sb.toString(); sb.setLength(0); i++; break }
                    ch == ',' && paren == 0 -> { args += sb.toString(); sb.setLength(0); i++ }
                    else -> { sb.append(ch); i++ }
                }
            }
            val nums = ArrayList<Double>()
            for (a in args) {
                val t = a.trim()
                if (t.isEmpty()) continue
                val range = rangeOf(t)
                if (range != null) {
                    for (r in range.top..range.bottom) for (cc in range.left..range.right) {
                        cellValue(r, cc)?.let { nums += it }
                    }
                } else {
                    Expr(m, t, depth + 1).parseAll()?.let { nums += it }
                }
            }
            return when (name) {
                "SUM" -> nums.sum()
                "AVERAGE", "AVG" -> if (nums.isEmpty()) null else nums.average()
                "MIN" -> nums.minOrNull()
                "MAX" -> nums.maxOrNull()
                "COUNT" -> nums.size.toDouble()
                "ABS" -> nums.firstOrNull()?.let { abs(it) }
                "ROUND" -> nums.firstOrNull()?.let { v ->
                    val digits = nums.getOrNull(1)?.toInt() ?: 0
                    BigDecimal(v).setScale(digits, java.math.RoundingMode.HALF_UP).toDouble()
                }
                else -> null
            }
        }
    }

    companion object {
        private const val KEY_STRIDE = 20000L
        private const val CHAR_PX = 7f
        private const val COL_PAD = 10f
        private const val MAX_PIXELS = 36_000_000f

        private fun key(row: Int, col: Int): Long = row.toLong() * KEY_STRIDE + col

        /** `AB12` → (row=11, col=27)，0 基。 */
        private fun refToCell(ref: String): Pair<Int, Int>? {
            if (ref.isEmpty()) return null
            var col = 0
            var i = 0
            while (i < ref.length && ref[i].isLetter()) {
                col = col * 26 + (ref[i].uppercaseChar() - 'A' + 1)
                i++
            }
            if (i == 0) return null
            val row = ref.substring(i).toIntOrNull() ?: return null
            return (row - 1) to (col - 1)
        }

        /** `A1:H32` → Rect(top,left,bottom,right) 0 基含端点。 */
        private fun rangeOf(ref: String): Rect? {
            val parts = ref.split(':')
            val a = refToCell(parts[0]) ?: return null
            val b = if (parts.size > 1) refToCell(parts[1]) ?: a else a
            return Rect(min(a.second, b.second), min(a.first, b.first), max(a.second, b.second), max(a.first, b.first))
        }
    }
}
