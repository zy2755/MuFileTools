package com.mutools.filetools.core

import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.Rect
import android.text.Layout
import android.text.SpannableString
import android.text.StaticLayout
import android.text.TextPaint
import android.text.style.AbsoluteSizeSpan
import android.text.style.ForegroundColorSpan
import android.text.style.StyleSpan
import android.text.style.UnderlineSpan
import java.math.BigDecimal
import java.math.MathContext
import java.text.DecimalFormat
import java.text.SimpleDateFormat
import java.util.Date
import java.util.TimeZone
import kotlin.math.abs
import kotlin.math.floor
import kotlin.math.max
import kotlin.math.roundToInt
import kotlin.math.sqrt

/**
 * 表格数据模型与绘制层。
 *
 * 单独抽出来是因为 .xlsx（zip + XML）和 .xls（复合文档 + BIFF8）除了"怎么读出数据"以外
 * 完全一样——读出来都是"一堆带样式索引的单元格 + 行高列宽 + 合并区域"，
 * 画法应该只有一份。解析器只负责把各自的格式翻译成 [SheetModel] + 样式表，
 * 剩下的几何计算、网格、边框、字号、数字格式都在这里。
 */
internal class SheetCell(
    val raw: String,
    val num: Double?,
    val isNumber: Boolean,
    val xf: Int,
    val formula: String? = null,
)

internal class SheetModel(
    val cells: HashMap<Long, SheetCell>,
    val rowH: HashMap<Int, Float>,       // 行 → 磅
    val colW: HashMap<Int, Float>,       // 列 → 字符宽
    val merges: List<Rect>,              // 合并区域（0 基，含端点）
    val maxRow: Int,
    val maxCol: Int,
    val defaultRowH: Float,
    val defaultColW: Float,
)

internal class TFont(val sz: Float, val bold: Boolean, val italic: Boolean, val color: Int?)

internal class TFill(val color: Int?)

internal class TBorder(val left: Int?, val right: Int?, val top: Int?, val bottom: Int?)

internal class TXf(
    val numFmtId: Int,
    val fontId: Int,
    val fillId: Int,
    val borderId: Int,
    val hAlign: String?,
    val vAlign: String?,
    val wrap: Boolean,
)

internal object SheetKey {
    const val STRIDE = 20000L

    fun of(row: Int, col: Int): Long = row.toLong() * STRIDE + col
    fun row(k: Long): Int = (k / STRIDE).toInt()
    fun col(k: Long): Int = (k % STRIDE).toInt()
}

/**
 * 把一个 [SheetModel] 画成位图。
 *
 * @param formula 公式兜底求值（文件里没存缓存值时才调），返回 null 表示算不出来
 */
internal class SheetPainter(
    private val fonts: List<TFont>,
    private val fills: List<TFill>,
    private val borders: List<TBorder>,
    private val xfs: List<TXf>,
    private val numFmts: Map<Int, String>,
    private val formula: ((SheetModel, SheetCell) -> Double?)? = null,
) {

    /** 版面几何，单位是 96 DPI 的像素。 */
    private class Geo(
        val colX: FloatArray, val colW: FloatArray,
        val rowY: FloatArray, val rowH: FloatArray,
        val width: Float, val height: Float,
    )

    /** 缩略图：宽固定，等比高。 */
    fun render(m: SheetModel, targetWidth: Int): Bitmap {
        val geo = geometry(m)
        val scale = (targetWidth.coerceAtLeast(32) / geo.width).coerceAtLeast(0.05f)
        return draw(m, geo, scale)
    }

    /** 导出：按 DPI 渲染，面积超限整体缩档防 OOM。 */
    fun renderDpi(m: SheetModel, dpi: Int): Bitmap {
        val geo = geometry(m)
        var scale = dpi / 96f
        val area = geo.width * scale * geo.height * scale
        if (area > MAX_PIXELS) scale *= sqrt(MAX_PIXELS / area)
        return draw(m, geo, scale)
    }

    // ---------------------------------------------------------------- 几何

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
            if (SheetKey.col(k) != col) continue
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
     * 粗略估宽：中文按全角、西文按约半个字符宽。
     * 宁可估宽一点——估窄了文字会被省略号截断，那是没法看的。
     */
    private fun estimateWidth(text: String): Float {
        var w = 0f
        for (ch in text.take(200)) {
            w += if (ch.code > 0x2E80) 17f else 8.5f
        }
        return w
    }

    /** 单元格要显示的文字（含数字/日期格式化）。 */
    private fun displayText(cell: SheetCell, m: SheetModel?): String {
        if (!cell.isNumber) return cell.raw
        // 缓存值优先；没有就现场算一遍公式（Excel 会重算，我们只做四则与常见聚合）
        val value = cell.num
            ?: cell.formula?.let { f -> m?.let { formula?.invoke(it, cell) } }
            ?: return ""
        val xf = xfs.getOrNull(cell.xf)
        val code = xf?.let { numFmts[it.numFmtId] ?: NumFmt.builtin(it.numFmtId) } ?: "General"
        return NumFmt.format(value, code)
    }

    // ---------------------------------------------------------------- 绘制

    private fun draw(m: SheetModel, geo: Geo, scale: Float): Bitmap {
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
            mergeByOrigin[SheetKey.of(r.top, r.left)] = r
            for (rr in r.top..r.bottom) for (cc in r.left..r.right) {
                if (rr != r.top || cc != r.left) covered += SheetKey.of(rr, cc)
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
            val row = SheetKey.row(k)
            val col = SheetKey.col(k)
            if (row < 0 || row >= geo.rowY.size - 1 || col < 0 || col >= geo.colX.size - 1) continue

            val merge = mergeByOrigin[k]
            val right = if (merge != null) merge.right.coerceAtMost(geo.colX.size - 2) else col
            val bottom = if (merge != null) merge.bottom.coerceAtMost(geo.rowY.size - 2) else row
            val rect = Rect(
                geo.colX[col].roundToInt(), geo.rowY[row].roundToInt(),
                geo.colX[right + 1].roundToInt(), geo.rowY[bottom + 1].roundToInt(),
            )
            if (rect.width() <= 0 || rect.height() <= 0) continue

            drawCell(c, rect, cell, xfs.getOrNull(cell.xf), borderPaint, m)
        }
        return bmp
    }

    private fun drawCell(
        c: Canvas, rect: Rect, cell: SheetCell, xf: TXf?, borderPaint: Paint, m: SheetModel,
    ) {
        val font = fonts.getOrNull(xf?.fontId ?: 0) ?: TFont(11f, false, false, null)
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

        val text = displayText(cell, m)
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
                        font.bold && font.italic -> android.graphics.Typeface.BOLD_ITALIC
                        font.bold -> android.graphics.Typeface.BOLD
                        else -> android.graphics.Typeface.ITALIC
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

    companion object {
        private const val CHAR_PX = 7f
        private const val COL_PAD = 10f

        /** 单表位图面积上限（约 6000×6000），超了整体缩档，防止低内存机 OOM。 */
        private const val MAX_PIXELS = 36_000_000f
    }
}

/** 数字格式：Excel 的格式串 → 显示文本。日期走 SimpleDateFormat，其余自己拼。 */internal object NumFmt {

    fun builtin(id: Int): String = when (id) {
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

    fun format(v: Double, code: String): String {
        // 公式求出 NaN / Infinity（除零、求值失败）时不能交给 BigDecimal —— 它会直接抛
        // NumberFormatException，把整页渲染带崩。Excel 在这种情况下显示错误串，这里留空。
        if (!v.isFinite()) return if (v.isNaN()) "" else "#NUM!"
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

    fun general(v: Double): String {
        if (v == floor(v) && abs(v) < 1e11) return v.toLong().toString()
        val bd = BigDecimal(v).round(MathContext(10)).stripTrailingZeros()
        return bd.toPlainString()
    }

    fun isDateFormat(code: String): Boolean {
        // 去掉引号里的字面量、方括号（颜色/条件）再判断
        var s = code.replace(Regex("\"[^\"]*\""), "").replace(Regex("\\[[^\\]]*\\]"), "")
        s = s.replace("General", "")
        return Regex("[yYmMdDhHsS]").containsMatchIn(s)
    }

    fun formatDate(serial: Double, code: String): String {
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
}

/**
 * 索引调色板（Excel 97 起的 56 色 + 系统色）。
 *
 * BIFF8 的 FONT.icv / XF 的边框前景色都是"调色板下标"，.xlsx 的 indexed 色也走同一套，
 * 所以放一份共用的。64/65 是"系统前景/背景"，按黑/白处理。
 */
internal object SheetPalette {

    private val COLORS = intArrayOf(
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
        0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
        0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
        0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
        0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
        0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
        0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
    )

    /** 调色板下标 → ARGB。未知下标退回黑色（Excel 的默认字色）。 */
    fun color(index: Int): Int = when {
        index in COLORS.indices -> (0xFF shl 24) or COLORS[index]
        index == 64 || index == 65 -> (0xFF shl 24)   // 系统前景/背景
        else -> (0xFF shl 24)
    }

    /** 边框里 0x7FFF 表示"没有边框颜色"，要跟黑色区分开。 */
    const val NO_COLOR = 0x7FFF
}
