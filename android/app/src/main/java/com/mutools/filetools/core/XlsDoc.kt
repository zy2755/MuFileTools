package com.mutools.filetools.core

import android.content.Context
import android.graphics.Bitmap
import android.graphics.Color
import android.graphics.Rect
import android.net.Uri
import java.io.ByteArrayOutputStream
import java.io.Closeable
import kotlin.math.abs

/**
 * XLS（Excel 97-2003，BIFF8）工作表渲染。
 *
 * 和 .xlsx 的区别只在"怎么把字节读成表"：.xls 是复合文档（[Cfb]）里的一条 Workbook 流，
 * 流里是一串 `type(2) + length(2) + payload` 的记录。读出来之后
 * 一样是 [SheetModel] + 样式表，绘制交给 [SheetPainter]（与 .xlsx 共用）。
 *
 * 解析要点都写在对应函数旁边，最容易踩的三个：
 *  1. 字符串是"共享字符串表（SST）"，还带 CONTINUE 记录接续，一条字符串可能横跨两条记录；
 *  2. 公式存的是 RPN 字节码（ptg 序列），不是文本；Excel 保存时会写缓存值，
 *     但 xlwt/openpyxl 这类库生成的文件没有，得自己算；
 *  3. 样式全在 XF 里按位域打包：填充、边框、对齐都挤在几个 32 位整数里。
 */
class XlsDoc(context: Context, uri: Uri) : PageSource {

    // ---------------------------------------------------------------- 中间态

    /** 解析中的单元格：公式此时还没算，等所有表都读完再统一求值（要跨表引用）。 */
    private class Pending(
        val xf: Int,
        val number: Boolean,
        val raw: String = "",
        val num: Double? = null,
        val tokens: ByteArray? = null,
    )

    private class SheetRef(val name: String, val bofPos: Int, val hidden: Boolean) {
        var pending: HashMap<Long, Pending>? = null
        var model: SheetModel? = null
    }

    // ---------------------------------------------------------------- 状态

    private val lock = Any()

    /** 整个文件读进内存：.xls 通常几十 KB ~ 几 MB，比维护随机读的流简单得多。 */
    private val bytes: ByteArray = context.contentResolver.openInputStream(uri)?.use {
        val buf = ByteArrayOutputStream(1 shl 16)
        it.copyTo(buf, 1 shl 16)
        buf.toByteArray()
    } ?: throw PageSourceException("打不开这个文件（可能已被删除或没有读取权限）")

    private val sheets = ArrayList<SheetRef>()

    private val fonts = ArrayList<TFont>()
    private val fills = ArrayList<TFill>()
    private val borders = ArrayList<TBorder>()
    private val xfs = ArrayList<TXf>()
    private val numFmts = HashMap<Int, String>()

    /** 外部表引用：EXTERNSHEET 表项 → 目标 sheet 序号，用来解 `Sheet2!A1` 这类 3D 引用。 */
    private val externSheets = ArrayList<Int>()

    private var date1904 = false
    private var defaultColW = 8f
    private var defaultRowH = 15f

    /**
     * 共享字符串表。**必须声明在 init 之前**：Kotlin 按声明顺序初始化属性，
     * 写在 init 后面的字段执行 init 时还是 null（而且 init 里赋的值随后会被自己的
     * 初始值覆盖回空）。
     */
    private var sst: List<String> = emptyList()

    /** 正在求值的单元格，防止循环引用把栈打穿。同样必须在 init 之前。 */
    private val evaluating = HashSet<Long>()

    /** 每张表的行高/列宽/合并区域，键是那张表的 cells（省得多一个包装类）。同样必须在 init 之前。 */
    private val pendingMeta =
        HashMap<HashMap<Long, Pending>, Triple<HashMap<Int, Float>, HashMap<Int, Float>, ArrayList<Rect>>>()

    init {
        val wb = readWorkbookStream()
        val recs = splitRecords(wb)
        parseGlobals(recs, wb)
        parseSheets(recs, wb)
    }

    override val pageCount: Int get() = sheets.size
    override val unitName: String get() = "个表"
    override fun label(index: Int): String = sheets.getOrNull(index)?.name ?: "表"
    override fun fileLabel(index: Int): String {
        val n = sheets.getOrNull(index)?.name ?: "表${index + 1}"
        return n.replace(Regex("[\\\\/:*?\"<>|]"), "_")
    }

    private val painter: SheetPainter by lazy { SheetPainter(fonts, fills, borders, xfs, numFmts) }

    override fun renderPage(index: Int, targetWidth: Int): Bitmap = synchronized(lock) {
        painter.render(modelOf(index), targetWidth)
    }

    override fun renderPageDpi(index: Int, dpi: Int): Bitmap = synchronized(lock) {
        painter.renderDpi(modelOf(index), dpi)
    }

    /** 数据全在内存里，没有需要释放的原生资源。 */
    override fun close() = Unit

    // ---------------------------------------------------------------- 容器

    /** 复合文档里取 Workbook 流（老文件叫 Book）。 */
    private fun readWorkbookStream(): ByteArray {
        val cfb = try {
            Cfb.open(bytes)
        } catch (e: Exception) {
            throw PageSourceException("这个 .xls 打不开：${e.message ?: "文件结构不完整"}")
        }
        return cfb.stream("Workbook")
            ?: cfb.stream("Book")
            ?: throw PageSourceException("这不是 Excel 文件（找不到 Workbook 数据流）")
    }

    /** 一条 BIFF 记录。 */
    private class Rec(val type: Int, val off: Int, val len: Int)

    /**
     * 扫描出所有记录的头。payload 不复制 —— SST 一条记录可能上百 KB，
     * 复制一遍纯属浪费，需要时按 off/len 从 wb 里取。
     */
    private fun splitRecords(wb: ByteArray): List<Rec> {
        val out = ArrayList<Rec>(4096)
        var i = 0
        while (i + 4 <= wb.size) {
            val type = Cfb.le16(wb, i)
            val len = Cfb.le16(wb, i + 2)
            if (type == 0 && len == 0) break          // 尾部补零，到底了
            if (i + 4 + len > wb.size) break
            out += Rec(type, i + 4, len)
            i += 4 + len
        }
        return out
    }

    // ---------------------------------------------------------------- 全局区

    private fun parseGlobals(recs: List<Rec>, wb: ByteArray) {
        var i = 0
        while (i < recs.size) {
            val r = recs[i]
            when (r.type) {
                // BOF 的记录类型永远是 0x0809，"是工作簿还是子表"写在负载里：
                // +0 vers(0x0600) / +2 dt / +4 rupBuild / +6 rupYear ……
                // 千万不要拿 dt 的值（0x0010 之类）去当记录类型比 —— 0x0040 本身就是一条真记录。
                BOF -> {
                    val dt = if (r.len >= 4) u16(wb, r.off + 2) else 0
                    if (dt != DT_WORKBOOK_GLOBALS) break
                }
                DATEMODE -> if (r.len >= 1) date1904 = (wb[r.off].toInt() and 0xFF) == 1
                // 字体索引 4 是 BIFF 的保留槽位（Excel 的默认字体占位），文件里没有对应的
                // FONT 记录。若不在这里补一个空位，从第 5 条字体起所有索引都会错一位 ——
                // 表现就是标题变白字（看不见）、表头变黑字。（xlrd 同样插了个 Dummy Font。）
                FONT -> {
                    if (fonts.size == 4) fonts += TFont(11f, false, false, null)
                    fonts += parseFont(wb, r)
                }
                FORMAT -> {
                    val id = u16(wb, r.off)
                    val code = readUnicode(wb, r.off + 2, r.off + r.len, lenLen = 2).first
                    numFmts[id] = code
                }
                XF -> xfs += parseXf(wb, r)
                BOUNDSHEET -> {
                    val bofPos = i32(wb, r.off)
                    val visibility = wb[r.off + 4].toInt() and 0xFF
                    val name = readUnicode(wb, r.off + 6, r.off + r.len, lenLen = 1).first
                    if (bofPos in 0 until wb.size) {
                        sheets += SheetRef(name.ifEmpty { "表${sheets.size + 1}" }, bofPos, visibility != 0)
                    }
                }
                SST -> {
                    // SST 后面可能跟若干条 CONTINUE，一条长字符串可能被切成两半
                    val chunks = ArrayList<ByteArray>()
                    chunks += wb.copyOfRange(r.off, r.off + r.len)
                    var j = i + 1
                    while (j < recs.size && recs[j].type == CONTINUE) {
                        chunks += wb.copyOfRange(recs[j].off, recs[j].off + recs[j].len)
                        j++
                    }
                    sst = parseSst(chunks)
                    i = j - 1
                }
                EXTERNSHEET -> {
                    // 记录布局：cXTI(2) + cXTI × 6 字节。**开头的 2 字节是条数，不是表项**，
                    // 漏掉它会整体错位 2 字节（所有 XTI 解析成 -1，跨表引用全失效）。
                    // 每个表项 = iSupBook(2) itabFirst(2) itabLast(2)，0=本工作簿，1+=外部；
                    // 真正要的是 itabFirst（第几张表）。
                    val n = if (r.len >= 2) u16(wb, r.off) else 0
                    for (k in 0 until minOf(n, (r.len - 2) / 6)) {
                        val base = r.off + 2 + k * 6
                        val sup = u16(wb, base)
                        val itab = u16(wb, base + 2)
                        externSheets += if (sup == 0 && itab < 0xFFFE) itab else -1
                    }
                }
                DEFCOLWIDTH -> if (r.len >= 2) {
                    val w = u16(wb, r.off)
                    if (w > 0) defaultColW = w.toFloat()
                }
                STANDARDWIDTH -> if (r.len >= 2) {
                    val w = u16(wb, r.off) / 256f
                    if (w > 0) defaultColW = w
                }
                DEFAULTROWHEIGHT -> if (r.len >= 4) {
                    val h = u16(wb, r.off + 2) / 20f
                    if (h > 0) defaultRowH = h
                }
            }
            i++
        }
    }

    // ---------------------------------------------------------------- 各子表

    private fun parseSheets(recs: List<Rec>, wb: ByteArray) {
        // 记录起始偏移 → 记录下标，方便按 BOUNDSHEET 给出的 BOF 位置跳过去
        val byOffset = HashMap<Int, Int>(recs.size * 2)
        recs.forEachIndexed { idx, r -> byOffset[r.off - 4] = idx }

        for (sheet in sheets) {
            val start = byOffset[sheet.bofPos] ?: continue
            sheet.pending = parseSheet(recs, start, wb)
        }
        // 所有表都读完了，公式才能跨表求值
        for (sheet in sheets) resolveSheet(sheet)
    }

    private fun parseSheet(recs: List<Rec>, start: Int, wb: ByteArray): HashMap<Long, Pending> {
        val cells = HashMap<Long, Pending>()
        val rowH = HashMap<Int, Float>()
        val colW = HashMap<Int, Float>()
        val merges = ArrayList<Rect>()
        var maxRow = 0
        var maxCol = 0
        var curRow = -1
        var lastFormula: Long = -1

        var i = start
        while (i < recs.size) {
            val r = recs[i]
            if (r.type == EOF) break
            when (r.type) {
                COLINFO -> if (r.len >= 10) {
                    val c1 = u16(wb, r.off)
                    val c2 = u16(wb, r.off + 2)
                    val width = u16(wb, r.off + 4) / 256f
                    val hidden = (u16(wb, r.off + 8) and 0x01) != 0
                    for (c in c1..minOf(c2, c1 + 255)) colW[c] = if (hidden) 0f else width
                    maxCol = maxOf(maxCol, c2)
                }
                ROW -> if (r.len >= 8) {
                    val row = u16(wb, r.off)
                    val h = u16(wb, r.off + 6) / 20f
                    if (h > 0f) rowH[row] = h
                    maxRow = maxOf(maxRow, row)
                }
                DIMENSIONS -> if (r.len >= 12) {
                    maxRow = maxOf(maxRow, i32(wb, r.off + 4) - 1)
                    maxCol = maxOf(maxCol, u16(wb, r.off + 10) - 1)
                }
                LABELSST -> if (r.len >= 10) {
                    val row = u16(wb, r.off); val col = u16(wb, r.off + 2)
                    val xf = u16(wb, r.off + 4); val idx = i32(wb, r.off + 6)
                    val text = sst.getOrElse(idx) { "" }
                    cells[SheetKey.of(row, col)] = Pending(xf, false, text)
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, col)
                }
                LABEL -> if (r.len >= 8) {
                    val row = u16(wb, r.off); val col = u16(wb, r.off + 2)
                    val xf = u16(wb, r.off + 4)
                    val text = readUnicode(wb, r.off + 6, r.off + r.len, lenLen = 2).first
                    cells[SheetKey.of(row, col)] = Pending(xf, false, text)
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, col)
                }
                NUMBER -> if (r.len >= 14) {
                    val row = u16(wb, r.off); val col = u16(wb, r.off + 2)
                    val xf = u16(wb, r.off + 4)
                    cells[SheetKey.of(row, col)] = Pending(xf, true, num = f64(wb, r.off + 6))
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, col)
                }
                RK -> if (r.len >= 10) {
                    val row = u16(wb, r.off); val col = u16(wb, r.off + 2)
                    val xf = u16(wb, r.off + 4)
                    cells[SheetKey.of(row, col)] = Pending(xf, true, num = decodeRk(i32(wb, r.off + 6)))
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, col)
                }
                MULRK -> if (r.len >= 6) {
                    val row = u16(wb, r.off)
                    val first = u16(wb, r.off + 2)
                    val n = (r.len - 6) / 6
                    for (k in 0 until n) {
                        val base = r.off + 4 + k * 6
                        val xf = u16(wb, base)
                        val v = decodeRk(i32(wb, base + 2))
                        cells[SheetKey.of(row, first + k)] = Pending(xf, true, num = v)
                    }
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, first + n - 1)
                }
                BLANK -> if (r.len >= 6) {
                    val row = u16(wb, r.off); val col = u16(wb, r.off + 2)
                    cells[SheetKey.of(row, col)] = Pending(u16(wb, r.off + 4), false)
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, col)
                }
                MULBLANK -> if (r.len >= 6) {
                    val row = u16(wb, r.off); val first = u16(wb, r.off + 2)
                    val n = (r.len - 6) / 2
                    for (k in 0 until n) {
                        val xf = u16(wb, r.off + 4 + k * 2)
                        // 空单元格也要留：它可能带底色或边框
                        cells.putIfAbsent(SheetKey.of(row, first + k), Pending(xf, false))
                    }
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, first + n - 1)
                }
                BOOLERR -> if (r.len >= 8) {
                    val row = u16(wb, r.off); val col = u16(wb, r.off + 2)
                    val xf = u16(wb, r.off + 4)
                    val v = wb[r.off + 6].toInt() and 0xFF
                    val isErr = (wb[r.off + 7].toInt() and 0xFF) == 1
                    val text = when {
                        isErr -> errorText(v)
                        v != 0 -> "TRUE"
                        else -> "FALSE"
                    }
                    cells[SheetKey.of(row, col)] = Pending(xf, false, text)
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, col)
                }
                FORMULA -> if (r.len >= 22) {
                    val row = u16(wb, r.off); val col = u16(wb, r.off + 2)
                    val xf = u16(wb, r.off + 4)
                    val key = SheetKey.of(row, col)
                    lastFormula = key
                    // 结果占 8 字节。要么是一整个 IEEE double，要么是"非数字结果"：
                    // 此时 **末尾两字节** 是 FF FF，开头第 1 字节写类型（0=字符串 1=布尔 2=错误 3=空）。
                    // 千万别把判据放在开头 —— xlwt/openpyxl 存的是 `03 00 … FF FF`，开头是 03。
                    val special = (wb[r.off + 12].toInt() and 0xFF) == 0xFF &&
                        (wb[r.off + 13].toInt() and 0xFF) == 0xFF
                    val kind = wb[r.off + 6].toInt() and 0xFF
                    val cce = u16(wb, r.off + 20)
                    val tokens = if (cce > 0 && r.off + 22 + cce <= r.off + r.len) {
                        wb.copyOfRange(r.off + 22, r.off + 22 + cce)
                    } else null
                    val pending = if (!special) {
                        Pending(xf, true, num = f64(wb, r.off + 6))
                    } else when (kind) {
                        // 0=字符串（紧跟一条 STRING 记录）、1=布尔、2=错误、3=空
                        1 -> Pending(xf, false, if ((wb[r.off + 8].toInt() and 0xFF) != 0) "TRUE" else "FALSE")
                        2 -> Pending(xf, false, errorText(wb[r.off + 8].toInt() and 0xFF))
                        else -> Pending(xf, true, tokens = tokens)
                    }
                    cells[key] = pending
                    maxRow = maxOf(maxRow, row); maxCol = maxOf(maxCol, col)
                }
                STRING -> if (r.len >= 3 && lastFormula >= 0) {
                    // 公式的字符串结果
                    val text = readUnicode(wb, r.off, r.off + r.len, lenLen = 2).first
                    cells[lastFormula] = Pending(cells[lastFormula]?.xf ?: 0, false, text)
                }
                MERGEDCELLS -> if (r.len >= 2) {
                    val n = u16(wb, r.off)
                    for (k in 0 until n) {
                        val base = r.off + 2 + k * 8
                        if (base + 8 > r.off + r.len) break
                        val r1 = u16(wb, base); val r2 = u16(wb, base + 2)
                        val c1 = u16(wb, base + 4); val c2 = u16(wb, base + 6)
                        merges += Rect(c1, r1, c2, r2)
                        maxRow = maxOf(maxRow, r2); maxCol = maxOf(maxCol, c2)
                    }
                }
            }
            if (r.type == ROW) curRow = u16(wb, r.off)
            i++
        }
        if (curRow >= 0) maxRow = maxOf(maxRow, curRow)

        return cells.also {
            pendingMeta[it] = Triple(rowH, colW, merges)
        }
    }

    // ---------------------------------------------------------------- 求值 + 建模

    private fun resolveSheet(sheet: SheetRef) {
        val pending = sheet.pending ?: return
        val meta = pendingMeta[pending] ?: return
        val cells = HashMap<Long, SheetCell>(pending.size * 2)
        for ((k, p) in pending) {
            val num = if (p.number && p.num == null && p.tokens != null) {
                valueAt(sheets.indexOf(sheet), SheetKey.row(k), SheetKey.col(k), 0)
            } else p.num
            cells[k] = SheetCell(p.raw, num, p.number, p.xf)
        }
        var maxRow = 0
        var maxCol = 0
        for (k in cells.keys) {
            maxRow = maxOf(maxRow, SheetKey.row(k))
            maxCol = maxOf(maxCol, SheetKey.col(k))
        }
        for (m in meta.third) {
            maxRow = maxOf(maxRow, m.bottom)
            maxCol = maxOf(maxCol, m.right)
        }
        sheet.model = SheetModel(
            cells, meta.first, meta.second, meta.third,
            maxRow, maxCol, defaultRowH, defaultColW,
        )
    }

    private fun modelOf(index: Int): SheetModel {
        val s = sheets[index]
        s.model?.let { return it }
        resolveSheet(s)
        return s.model ?: SheetModel(
            HashMap(), HashMap(), HashMap(), emptyList(), 0, 0, defaultRowH, defaultColW,
        )
    }

    /** 正在求值的单元格，防止 A1→B1→A1 这种死循环。 */

    private fun valueAt(sheetIdx: Int, row: Int, col: Int, depth: Int): Double? {
        if (depth > 8) return null
        val sheet = sheets.getOrNull(sheetIdx) ?: return null
        val pending = sheet.pending ?: return null
        val k = SheetKey.of(row, col)
        val cell = pending[k] ?: return null
        cell.num?.let { return it }
        if (!cell.number) return null
        val tokens = cell.tokens ?: return null
        val guard = sheetIdx.toLong() * 1_000_000_000L + k
        if (!evaluating.add(guard)) return null
        return try {
            PtgEval(sheetIdx, depth).eval(tokens)
        } finally {
            evaluating.remove(guard)
        }
    }

    // ---------------------------------------------------------------- BIFF 工具

    /** RK 编码：最低位=÷100，次低位=整数(30 位有符号)还是 double 高 30 位。 */
    private fun decodeRk(rk: Int): Double {
        val cents = (rk and 0x01) != 0
        val isInt = (rk and 0x02) != 0
        var v = if (isInt) {
            (rk shr 2).toDouble()
        } else {
            java.lang.Double.longBitsToDouble((rk.toLong() and 0xFFFFFFFCL) shl 32)
        }
        if (cents) v /= 100.0
        return v
    }

    private fun errorText(code: Int): String = when (code) {
        0x00 -> "#NULL!"
        0x07 -> "#DIV/0!"
        0x0F -> "#VALUE!"
        0x17 -> "#REF!"
        0x1D -> "#NAME?"
        0x24 -> "#NUM!"
        0x2A -> "#N/A"
        else -> "#ERR"
    }

    private fun u16(b: ByteArray, off: Int): Int = Cfb.le16(b, off)
    private fun i32(b: ByteArray, off: Int): Int = Cfb.le32(b, off)
    private fun f64(b: ByteArray, off: Int): Double =
        java.lang.Double.longBitsToDouble(
            (Cfb.le32(b, off).toLong() and 0xFFFFFFFFL) or
                (Cfb.le32(b, off + 4).toLong() shl 32)
        )

    /**
     * BIFF8 的 unicode 字符串：`cch(1或2) + flags(1) + [rt(2)] + [phonetic(4)] + chars`。
     * flags 最低位 =1 是 UTF-16，0 是"压缩"的 latin-1（注意不是 ASCII）。
     */
    private fun readUnicode(b: ByteArray, pos: Int, end: Int, lenLen: Int): Pair<String, Int> {
        if (pos + lenLen > end) return "" to end
        val cch = if (lenLen == 1) (b[pos].toInt() and 0xFF) else u16(b, pos)
        var p = pos + lenLen
        if (p >= end) return "" to end
        val flags = b[p].toInt() and 0xFF
        p++
        var rt = 0
        if (flags and 0x08 != 0) {
            rt = u16(b, p); p += 2
        }
        if (flags and 0x04 != 0) p += 4
        val wide = flags and 0x01 != 0
        val sb = StringBuilder(cch)
        if (wide) {
            for (k in 0 until cch) {
                if (p + 2 > end) break
                sb.append(u16(b, p).toChar())
                p += 2
            }
        } else {
            for (k in 0 until cch) {
                if (p >= end) break
                sb.append((b[p].toInt() and 0xFF).toChar())
                p++
            }
        }
        p += rt * 4
        return sb.toString() to p
    }

    /**
     * 共享字符串表。一条字符串可能横跨 SST 与 CONTINUE：
     * 跨记录时**新记录的第一个字节是新的 flags**（编码可能在 UTF-16/压缩之间切换），
     * 这一点漏掉会把后半截读成乱码。
     */
    private fun parseSst(chunks: List<ByteArray>): List<String> {
        if (chunks.isEmpty()) return emptyList()
        var ci = 0
        var b = chunks[0]
        var pos = 8                                     // 跳过 cstTotal / cstUnique
        val out = ArrayList<String>(u16(b, 4))

        fun nextChunk(): Boolean {
            if (ci + 1 >= chunks.size) return false
            ci++
            b = chunks[ci]
            pos = 0
            return true
        }

        while (true) {
            if (pos + 3 > b.size) { if (!nextChunk()) break else continue }
            val cch = u16(b, pos); pos += 2
            var flags = b[pos].toInt() and 0xFF; pos++
            var rt = 0
            if (flags and 0x08 != 0) { rt = u16(b, pos); pos += 2 }
            var phonetic = 0
            if (flags and 0x04 != 0) { phonetic = i32(b, pos); pos += 4 }

            val sb = StringBuilder(cch)
            var got = 0
            while (got < cch) {
                val wide = flags and 0x01 != 0
                if (wide) {
                    val avail = (b.size - pos) / 2
                    if (avail <= 0) {
                        if (!nextChunk()) return out
                        flags = b[pos].toInt() and 0xFF; pos++
                        continue
                    }
                    val take = minOf(avail, cch - got)
                    for (k in 0 until take) {
                        sb.append(u16(b, pos + k * 2).toChar())
                    }
                    pos += take * 2
                    got += take
                } else {
                    val avail = b.size - pos
                    if (avail <= 0) {
                        if (!nextChunk()) return out
                        flags = b[pos].toInt() and 0xFF; pos++
                        continue
                    }
                    val take = minOf(avail, cch - got)
                    for (k in 0 until take) sb.append((b[pos + k].toInt() and 0xFF).toChar())
                    pos += take
                    got += take
                }
                if (got < cch && !nextChunk()) return out
                if (got < cch) { flags = b[pos].toInt() and 0xFF; pos++ }
            }

            pos += rt * 4
            pos += phonetic

            out += sb.toString()
            // 记录刚好读完 → 换下一条（可能还有 SST 之后的 CONTINUE）
            if (pos >= b.size) {
                if (!nextChunk()) break
            }
        }
        return out
    }

    // ---------------------------------------------------------------- 样式

    private fun parseFont(wb: ByteArray, r: Rec): TFont {
        if (r.len < 16) return TFont(11f, false, false, null)
        val heightTwips = u16(wb, r.off)
        val grbit = u16(wb, r.off + 2)
        val icv = u16(wb, r.off + 4)
        val weight = u16(wb, r.off + 6)
        val sz = if (heightTwips > 0) heightTwips / 20f else 11f
        val bold = weight >= 700 || (grbit and 0x0001) != 0
        val italic = (grbit and 0x0002) != 0
        val color = if (icv == 0x7FFF) null else SheetPalette.color(icv)
        return TFont(sz, bold, italic, color)
    }

    private fun parseXf(wb: ByteArray, r: Rec): TXf {
        if (r.len < 16) return TXf(0, 0, 0, 0, null, null, false)
        val ifnt = u16(wb, r.off)
        val ifmt = u16(wb, r.off + 2)
        val alc1 = wb[r.off + 6].toInt() and 0xFF
        val alc2 = wb[r.off + 8].toInt() and 0xFF
        val b1 = i32(wb, r.off + 10)
        val b2 = if (r.len >= 18) i32(wb, r.off + 14) else 0
        val b3 = if (r.len >= 20) u16(wb, r.off + 18) else 0

        val hAlign = when (alc1 and 0x07) {
            1 -> "left"
            2 -> "center"
            3 -> "right"
            6 -> "centerContinuous"
            else -> null
        }
        val vAlign = when ((alc1 ushr 4) and 0x07) {
            0 -> "top"
            1 -> "center"
            2 -> "bottom"
            else -> null
        }
        val wrap = (alc1 and 0x08) != 0
        val indent = alc2 and 0x0F
        val useIndent = indent > 0 && hAlign == "left"

        // 边框：4 条线各 4 位样式（0=无线），颜色分别是 7 位调色板下标
        fun line(style: Int, colorIdx: Int): Int? =
            if (style == 0) null else {
                val c = if (colorIdx == 0x7F || colorIdx == 0x40) 0 else colorIdx
                SheetPalette.color(c)
            }

        val leftStyle = b1 and 0x0F
        val rightStyle = (b1 ushr 4) and 0x0F
        val topStyle = (b1 ushr 8) and 0x0F
        val bottomStyle = (b1 ushr 12) and 0x0F
        val leftColor = (b1 ushr 16) and 0x7F
        val rightColor = (b1 ushr 23) and 0x3F
        val topColor = b2 and 0x7F
        val bottomColor = (b2 ushr 7) and 0x3F

        val bd = TBorder(
            if (leftStyle == 0) null else line(leftStyle, leftColor) ?: 0xFF808080.toInt(),
            if (rightStyle == 0) null else line(rightStyle, rightColor) ?: 0xFF808080.toInt(),
            if (topStyle == 0) null else line(topStyle, topColor) ?: 0xFF808080.toInt(),
            if (bottomStyle == 0) null else line(bottomStyle, bottomColor) ?: 0xFF808080.toInt(),
        )

        // 填充：pattern 0=无，1=实心（前景色），2..18 是各种网纹，退化成前景色
        val pattern = (b2 ushr 26) and 0x3F
        val patternColor = b3 and 0x7F
        val fill = if (pattern == 0) TFill(null)
        else TFill(if (patternColor == 0x40 || patternColor == 0x41) null else SheetPalette.color(patternColor))

        return TXf(
            numFmtId = ifmt,
            fontId = ifnt,
            fillId = addFill(fill),
            borderId = addBorder(bd),
            hAlign = if (useIndent) "left" else hAlign,
            vAlign = vAlign,
            wrap = wrap,
        )
    }

    /** 样式表去重：BIFF 的填充/边框是内联在 XF 里的，转成"索引"才能喂给公共绘制层。 */
    private fun addFill(f: TFill): Int {
        val i = fills.indexOfFirst { it.color == f.color }
        if (i >= 0) return i
        fills += f
        return fills.size - 1
    }

    private fun addBorder(b: TBorder): Int {
        val i = borders.indexOfFirst {
            it.left == b.left && it.right == b.right && it.top == b.top && it.bottom == b.bottom
        }
        if (i >= 0) return i
        borders += b
        return borders.size - 1
    }

    // ---------------------------------------------------------------- 公式（ptg 字节码）

    /**
     * BIFF 公式存的是 RPN 字节码，不是文本：`44 0200 02C0` 就是"引用 C3"，
     * `1F <8字节double>` 是数字常量，`42 01 0500` 是"取 1 个参数的 5 号函数(AVERAGE)"。
     * 这里只做四则、百分号、括号、引用/区域、以及常见的聚合函数，
     * 遇到不认识的 token 就放弃（返回 null，单元格留空，不瞎猜）。
     */
    private inner class PtgEval(private val sheetIdx: Int, private val depth: Int) {

        private val stack = ArrayList<Any>(16)

        fun eval(tok: ByteArray): Double? {
            var p = 0
            while (p < tok.size) {
                val op = tok[p].toInt() and 0xFF
                p++
                when (op) {
                    0x03 -> binary { a, b -> a + b }
                    0x04 -> binary { a, b -> a - b }
                    0x05 -> binary { a, b -> a * b }
                    0x06 -> binary { a, b ->
                        if (b == 0.0) return@binary Double.NaN else a / b
                    }
                    0x07 -> binary { a, b -> Math.pow(a, b) }
                    0x12 -> { /* 一元加，啥也不用做 */ }
                    0x13 -> { push(-(popNum() ?: return null)) }
                    0x14 -> { push((popNum() ?: return null) / 100.0) }
                    0x15 -> { /* 括号，无操作 */ }
                    0x16 -> { p += 0 }                              // 缺省参数
                    0x17 -> {                                       // 字符串常量
                        val cch = tok[p].toInt() and 0xFF; p++
                        val flags = tok[p].toInt() and 0xFF; p++
                        if (flags and 0x01 != 0) {
                            val sb = StringBuilder(cch)
                            for (k in 0 until cch) { sb.append(u16(tok, p).toChar()); p += 2 }
                            push(sb.toString())
                        } else {
                            val sb = StringBuilder(cch)
                            for (k in 0 until cch) { sb.append((tok[p].toInt() and 0xFF).toChar()); p++ }
                            push(sb.toString())
                        }
                    }
                    0x19 -> p += 3                                  // ptgAttr（含 SUM 的优化标记）
                    0x1C -> { p += 1; push(Double.NaN) }             // 错误值
                    0x1D -> { push(if ((tok[p].toInt() and 0xFF) != 0) 1.0 else 0.0); p += 1 }
                    0x1E -> { push(u16(tok, p).toDouble()); p += 2 } // 16 位整数
                    0x1F -> { push(f64(tok, p)); p += 8 }           // double
                    0x24, 0x44, 0x64, 0x2C -> {                     // 单元格引用
                        push(refValue(u16(tok, p), u16(tok, p + 2)) ?: return null)
                        p += 4
                    }
                    0x25, 0x45, 0x65, 0x2D -> {                     // 区域
                        // ptgArea: rwFirst(2) rwLast(2) colFirst(2) colLast(2)
                        push(RangeRef(
                            u16(tok, p), u16(tok, p + 4),
                            u16(tok, p + 2), u16(tok, p + 6),
                        ))
                        p += 8
                    }
                    0x3A, 0x5A, 0x7A -> {                           // 3D 引用（跨表）
                        val ixti = u16(tok, p)
                        val s = sheetOf(ixti) ?: return null
                        push(refValueIn(s, u16(tok, p + 2), u16(tok, p + 4)) ?: return null)
                        p += 6
                    }
                    0x3B, 0x5B, 0x7B -> {                           // 3D 区域
                        val ixti = u16(tok, p)
                        val s = sheetOf(ixti) ?: return null
                        push(RangeRefIn(
                            s,
                            u16(tok, p + 2), u16(tok, p + 6),
                            u16(tok, p + 4), u16(tok, p + 8),
                        ))
                        p += 10
                    }
                    0x21, 0x41, 0x61 -> {
                        val ftab = u16(tok, p); p += 2
                        if (!callFunc(ftab, 0)) return null
                    }
                    0x22, 0x42, 0x62 -> {
                        val cparams = tok[p].toInt() and 0xFF; p++
                        val ftab = u16(tok, p); p += 2
                        if (!callFunc(ftab, cparams)) return null
                    }
                    else -> return null                              // 不认识就放弃
                }
            }
            return (stack.lastOrNull() as? Double)
        }

        private fun push(v: Any) { stack += v }

        private fun popNum(): Double? {
            if (stack.isEmpty()) return null
            val v = stack.removeAt(stack.size - 1)
            return when (v) {
                is Double -> v
                is RangeRef -> null
                is RangeRefIn -> null
                else -> null
            }
        }

        private inline fun binary(f: (Double, Double) -> Double) {
            if (stack.size < 2) { stack.clear(); return }
            val b = stack.removeAt(stack.size - 1)
            val a = stack.removeAt(stack.size - 1)
            val x = a as? Double ?: return
            val y = b as? Double ?: return
            stack += f(x, y)
        }

        private fun rowOf(v: Int): Int = v
        private fun colOf(v: Int): Int = v and 0x3FFF

        private fun refValue(row: Int, colFlags: Int): Double? =
            valueAt(sheetIdx, row, colOf(colFlags), depth + 1)

        private fun refValueIn(s: Int, row: Int, colFlags: Int): Double? =
            valueAt(s, row, colOf(colFlags), depth + 1)

        /** ixti → 本工作簿内的表序号（EXTERNSHEET 里记的是 supbook 内的 itab）。 */
        private fun sheetOf(ixti: Int): Int? = externSheets.getOrNull(ixti)?.takeIf { it in sheets.indices }

        private fun numbers(v: Any): List<Double> = when (v) {
            is Double -> listOf(v)
            is RangeRef -> rangeValues(sheetIdx, v.r1, v.c1, v.r2, v.c2, depth)
            is RangeRefIn -> rangeValues(v.sheet, v.r1, v.c1, v.r2, v.c2, depth)
            else -> emptyList()
        }

        private fun callFunc(ftab: Int, cparams: Int): Boolean {
            val args = ArrayList<Any>(cparams)
            for (k in 0 until cparams) {
                if (stack.isEmpty()) return false
                args.add(0, stack.removeAt(stack.size - 1))
            }
            when (ftab) {
                0 -> stack += numbers(args).size.toDouble()                  // COUNT
                4 -> stack += numbers(args).sum()                            // SUM
                5 -> {                                                       // AVERAGE
                    val v = numbers(args)
                    if (v.isEmpty()) return false
                    stack += v.sum() / v.size
                }
                6 -> {                                                       // MIN
                    val v = numbers(args)
                    if (v.isEmpty()) return false
                    stack += v.min()
                }
                7 -> {                                                       // MAX
                    val v = numbers(args)
                    if (v.isEmpty()) return false
                    stack += v.max()
                }
                24 -> {                                                      // ABS
                    val a = args.getOrNull(0) as? Double ?: return false
                    stack += abs(a)
                }
                25 -> {                                                      // INT
                    val a = args.getOrNull(0) as? Double ?: return false
                    stack += kotlin.math.floor(a)
                }
                27 -> {                                                      // ROUND
                    val a = args.getOrNull(0) as? Double ?: return false
                    val d = (args.getOrNull(1) as? Double ?: 0.0).toInt()
                    stack += BigDecimalOf(a).setScale(d, java.math.RoundingMode.HALF_UP).toDouble()
                }
                39 -> {                                                      // MOD
                    val a = args.getOrNull(0) as? Double ?: return false
                    val b = args.getOrNull(1) as? Double ?: return false
                    if (b == 0.0) return false
                    stack += a - kotlin.math.floor(a / b) * b
                }
                169 -> {                                                     // COUNTA
                    var n = 0
                    for (a in args) {
                        when (a) {
                            is Double -> n++
                            is String -> if (a.isNotEmpty()) n++
                            is RangeRef -> n += rangeCountNonEmpty(sheetIdx, a.r1, a.c1, a.r2, a.c2)
                            is RangeRefIn -> n += rangeCountNonEmpty(a.sheet, a.r1, a.c1, a.r2, a.c2)
                        }
                    }
                    stack += n.toDouble()
                }
                else -> return false
            }
            return true
        }

        private fun numbers(args: List<Any>): List<Double> {
            val out = ArrayList<Double>()
            for (a in args) out += numbers(a)
            return out
        }

        private fun BigDecimalOf(v: Double) = java.math.BigDecimal.valueOf(v)
    }

    /** 区域引用。字段顺序统一为 (r1, c1, r2, c2) —— ptgArea 存的是 rwFirst/rwLast/colFirst/colLast，
     *  这里换过一次序，注释写清楚免得再错。 */
    private class RangeRef(val r1: Int, val c1: Int, val r2: Int, val c2: Int)

    private class RangeRefIn(
        val sheet: Int, val r1: Int, val c1: Int, val r2: Int, val c2: Int,
    )

    private fun rangeValues(sheet: Int, r1: Int, c1: Int, r2: Int, c2: Int, depth: Int = 0): List<Double> {
        val out = ArrayList<Double>()
        val s = sheets.getOrNull(sheet)?.pending ?: return out
        val rFrom = maxOf(0, minOf(r1, r2))
        val rTo = maxOf(r1, r2)
        val cFrom = maxOf(0, minOf(c1 and 0x3FFF, c2 and 0x3FFF))
        val cTo = maxOf(c1 and 0x3FFF, c2 and 0x3FFF)
        if (rTo - rFrom > 5000 || cTo - cFrom > 500) return out     // 防呆
        for (r in rFrom..rTo) for (c in cFrom..cTo) {
            val cell = s[SheetKey.of(r, c)] ?: continue
            if (!cell.number) continue
            // 区域里也可能有"有公式没缓存值"的格子（AVERAGE 里套总评列就是），要现场算
            out += cell.num ?: valueAt(sheet, r, c, depth + 1) ?: continue
        }
        return out
    }

    private fun rangeCountNonEmpty(sheet: Int, r1: Int, c1: Int, r2: Int, c2: Int): Int {
        val s = sheets.getOrNull(sheet)?.pending ?: return 0
        val rFrom = maxOf(0, minOf(r1, r2))
        val rTo = maxOf(r1, r2)
        val cFrom = maxOf(0, minOf(c1 and 0x3FFF, c2 and 0x3FFF))
        val cTo = maxOf(c1 and 0x3FFF, c2 and 0x3FFF)
        if (rTo - rFrom > 5000 || cTo - cFrom > 500) return 0
        var n = 0
        for (r in rFrom..rTo) for (c in cFrom..cTo) {
            val cell = s[SheetKey.of(r, c)] ?: continue
            if (cell.number || cell.raw.isNotEmpty()) n++
        }
        return n
    }

    companion object {
        // 记录类型
        private const val EOF = 0x000A
        private const val BOUNDSHEET = 0x0085
        private const val CONTINUE = 0x003C
        private const val SST = 0x00FC
        private const val FONT = 0x0031
        private const val FORMAT = 0x041E
        private const val XF = 0x00E0
        private const val DATEMODE = 0x0022
        private const val EXTERNSHEET = 0x0017
        private const val DEFCOLWIDTH = 0x0055
        private const val STANDARDWIDTH = 0x0099
        private const val DEFAULTROWHEIGHT = 0x0225
        private const val COLINFO = 0x007D
        private const val ROW = 0x0208
        private const val DIMENSIONS = 0x0200
        private const val LABELSST = 0x00FD
        private const val LABEL = 0x0204
        private const val NUMBER = 0x0203
        private const val RK = 0x027E
        private const val MULRK = 0x00BD
        private const val BLANK = 0x0201
        private const val MULBLANK = 0x00BE
        private const val BOOLERR = 0x0205
        private const val FORMULA = 0x0006
        private const val STRING = 0x0207
        private const val MERGEDCELLS = 0x00E5

        private const val BOF = 0x0809

        /** BOF 负载里的 dt：说明这个子流是工作簿全局区、工作表、图表还是宏表。 */
        private const val DT_WORKBOOK_GLOBALS = 0x0005
    }
}
