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
    // 单元格/样式/绘制的公共部分在 SheetTable.kt（.xls 那边也用同一份）

    private class SheetRef(val name: String, val part: String) {
        var model: SheetModel? = null
    }

    // ---------------------------------------------------------------- 状态

    private val lock = Any()
    private val opc = Opc.open(context, uri, ".xlsx")
    private val mainPart: String = opc.firstByType("", "officeDocument") ?: "xl/workbook.xml"

    private val sheets = ArrayList<SheetRef>()
    private var sharedStrings: List<String> = emptyList()

    private val fonts = ArrayList<TFont>()
    private val fills = ArrayList<TFill>()
    private val borders = ArrayList<TBorder>()
    private val xfs = ArrayList<TXf>()
    private val numFmts = HashMap<Int, String>()
    private val themeColors = HashMap<Int, Int>()

    init {
        parseTheme()
        parseSharedStrings()
        parseStyles()
        parseWorkbook()
    }

    /** 绘制层：样式表在 parseStyles 之后才稳定，所以用 lazy。 */
    private val painter: SheetPainter by lazy {
        SheetPainter(fonts, fills, borders, xfs, numFmts) { m, cell ->
            cell.formula?.let { evalFormula(m, it, 0) }
        }
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

    private fun parseFont(p: XmlPullParser): TFont {
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
        return TFont(sz, bold, italic, color)
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

    /** 调色板下标（老文件会用 indexed；与 .xls 的 BIFF 调色板是同一张表） */
    private fun indexedColor(i: Int): Int = SheetPalette.color(i)

    private fun parseFill(p: XmlPullParser): TFill {
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
        return TFill(color)
    }

    private fun parseBorder(p: XmlPullParser): TBorder {
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
        return TBorder(l, r, t, b)
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

    private fun parseXf(p: XmlPullParser): TXf {
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
        return TXf(numFmtId, fontId, fillId, borderId, hAlign, vAlign, wrap)
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
        val cells = HashMap<Long, SheetCell>()
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

    private fun parseCell(p: XmlPullParser): SheetCell? {
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
                SheetCell(sharedStrings.getOrElse(idx) { "" }, null, false, s)
            }
            "inlineStr" -> if (inline.isEmpty()) null else SheetCell(inline, null, false, s)
            "str" -> if (v.isEmpty()) null else SheetCell(v, null, false, s)
            "b" -> SheetCell(if (v.trim() == "1") "TRUE" else "FALSE", null, false, s)
            "e" -> SheetCell(v, null, false, s)
            "d" -> SheetCell(v, null, false, s)
            else -> {
                val d = v.trim().toDoubleOrNull()
                // 只有公式、没有缓存值（openpyxl 等生成的文件很常见）——先留着，渲染时再算
                if (d == null && formula != null) SheetCell("", null, true, s, formula)
                else if (d == null) null
                else SheetCell(v, d, true, s, formula)
            }
        }
    }

    // ---------------------------------------------------------------- 渲染
    // 几何计算、网格、边框、字号、数字格式都在 SheetTable.kt（.xls 共用同一份）

    override fun renderPage(index: Int, targetWidth: Int): Bitmap = synchronized(lock) {
        painter.render(modelOf(index), targetWidth)
    }

    override fun renderPageDpi(index: Int, dpi: Int): Bitmap = synchronized(lock) {
        painter.renderDpi(modelOf(index), dpi)
    }

    override fun close() {
        opc.close()
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
        private fun key(row: Int, col: Int): Long = SheetKey.of(row, col)

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
