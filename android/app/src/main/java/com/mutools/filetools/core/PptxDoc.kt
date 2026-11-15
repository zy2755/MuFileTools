package com.mutools.filetools.core

import android.content.Context
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.Path
import android.graphics.Rect
import android.graphics.RectF
import android.graphics.Typeface
import android.net.Uri
import android.text.Layout
import android.text.SpannableString
import android.text.StaticLayout
import android.text.TextPaint
import android.text.style.AbsoluteSizeSpan
import android.text.style.ForegroundColorSpan
import android.text.style.StyleSpan
import android.text.style.UnderlineSpan
import android.util.LruCache
import org.xmlpull.v1.XmlPullParser
import kotlin.math.max
import kotlin.math.min

/**
 * PPTX 幻灯片渲染（纯 Kotlin，不依赖 Office / native 引擎）。
 *
 * 结构照 WPS 的做法：presentation.xml 定页序与画布 → 逐页 slideN.xml 取 spTree →
 * 形状再按 <a:xfrm> 画到 Canvas 上。和 Word 不同，PPT 天生就是"固定版式"，
 * 每页有自己的坐标系，所以**不需要排版引擎**，直接按坐标画就基本一致。
 *
 * 三个必须处理的继承（否则样例这类文件会整页空白）：
 *   1. 占位符（p:ph）的坐标常在 slideLayout / slideMaster 里，slide 自身 spPr 是空的；
 *   2. 字号/颜色/对齐也在版式里（lstStyle / txStyles）；
 *   3. 背景色多在母版（p:bgRef 引用主题的 bgFillStyleLst）。
 *
 * 坐标系：全程用 EMU（1 cm = 360000，1 pt = 12700）画，最后整体 scale 到像素，
 * 这样字号、线宽、间距只写一次，导出的 DPI 只是缩放系数。
 */
class PptxDoc(context: Context, uri: Uri) : PageSource {

    // ---------------------------------------------------------------- 模型

    private enum class Kind { SP, PIC, TABLE, LINE, GROUP, FRAME }

    private class Xfrm(
        var x: Float = 0f, var y: Float = 0f, var cx: Float = 0f, var cy: Float = 0f,
        var chOffX: Float = 0f, var chOffY: Float = 0f, var chExtX: Float = 0f, var chExtY: Float = 0f,
        var rotDeg: Float = 0f, var flipH: Boolean = false, var flipV: Boolean = false,
    )

    private class Fill(val none: Boolean = false, val solid: Int? = null, val imagePart: String? = null)

    private class Line(var color: Int = Color.BLACK, var widthEmu: Float = 12700f, var none: Boolean = false)

    private class Run(
        val text: String,
        val sz: Int? = null,          // 百分之一磅
        val bold: Boolean? = null,
        val italic: Boolean? = null,
        val underline: Boolean? = null,
        val color: Int? = null,
    )

    private class Para(
        val runs: List<Run>,
        val algn: String? = null,
        val marL: Float? = null,
        val indent: Float? = null,
        val lvl: Int = 0,
        val bullet: String? = null,
        val linePct: Float? = null,     // 百分比（100 = 单倍）
        val spaceBeforePt: Float? = null,
    )

    private class Body(
        val paras: List<Para>,
        val anchor: String? = null,
        val lIns: Float = 91440f, val tIns: Float = 45720f,
        val rIns: Float = 91440f, val bIns: Float = 45720f,
        val wrap: Boolean = true,
        val lstStyle: TxtStyle? = null,
    )

    /** 一档段落默认样式（对应 lvlXpPr）。字段为 null 表示"没写，继续往上找"。 */
    private class LvlStyle(
        val sz: Int? = null,
        val bold: Boolean? = null,
        val italic: Boolean? = null,
        val color: Int? = null,
        val algn: String? = null,
        val marL: Float? = null,
        val bullet: String? = null,
    ) {
        fun merge(o: LvlStyle?): LvlStyle {
            if (o == null) return this
            return LvlStyle(
                sz ?: o.sz, bold ?: o.bold, italic ?: o.italic, color ?: o.color,
                algn ?: o.algn, marL ?: o.marL, bullet ?: o.bullet,
            )
        }
    }

    private class TxtStyle(val lvls: Array<LvlStyle?> = arrayOfNulls(9)) {
        operator fun get(lvl: Int): LvlStyle? = lvls[(lvl - 1).coerceIn(0, 8)]
    }

    private class Cell(
        val text: Body?,
        val fill: Int? = null,
        val colSpan: Int = 1, val rowSpan: Int = 1,
        val anchor: String? = null,
    )

    private class Table(
        val colW: FloatArray,
        val rows: List<List<Cell>>,
        val rowH: FloatArray,
    )

    private class PhStyle(val fillRef: Int? = null, val lineRef: Int? = null, val fontRef: Int? = null)

    private class Node(val kind: Kind) {
        var phType: String? = null
        var phIdx: Int? = null
        var xfrm: Xfrm? = null
        var geom: String = "rect"
        var fill: Fill? = null
        var line: Line? = null
        var body: Body? = null
        var image: String? = null
        var srcRect: RectF? = null
        var table: Table? = null
        var style: PhStyle? = null
        var children: List<Node>? = null
        var chartPart: String? = null
    }

    private class Slide(val part: String) {
        var shapes: List<Node>? = null
        var layout: String? = null
        var layoutLoaded = false
        var bg: Fill? = null
    }

    // ---------------------------------------------------------------- 状态

    private val lock = Any()

    /** 当前这一帧的「像素 / EMU」，render() 每次刷新；文字排版要换算成像素才不会失真。 */
    private var pxPerEmu = 1f
    private val opc = Opc.open(context, uri, ".pptx")
    private val mainPart: String =
        opc.firstByType("", "officeDocument") ?: "ppt/presentation.xml"

    private val slides = ArrayList<Slide>()
    private var slideCx = 12192000f
    private var slideCy = 6858000f

    private var masterPart: String? = null
    private var masterShapes: List<Node> = emptyList()
    private var masterBg: Fill? = null
    private var masterTitle: TxtStyle? = null
    private var masterBody: TxtStyle? = null
    private var masterOther: TxtStyle? = null
    private var defaultText: TxtStyle? = null

    private val layoutShapesCache = HashMap<String, List<Node>>()
    private val layoutBgCache = HashMap<String, Fill?>()

    private val scheme = HashMap<String, Int>()
    private val clrMap = HashMap<String, String>()

    private val imgCache = object : LruCache<String, Bitmap>(16 * 1024 * 1024) {
        override fun sizeOf(key: String, value: Bitmap): Int = value.byteCount
    }

    init {
        parsePresentation()
        parseMaster()
    }

    override val pageCount: Int get() = slides.size
    override val unitName: String get() = "页"
    override fun label(index: Int): String = "第 ${index + 1} 页"
    override fun fileLabel(index: Int): String = (index + 1).toString()

    // ---------------------------------------------------------------- 打开：presentation.xml

    private fun parsePresentation() {
        val p = opc.parser(mainPart) ?: return
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.eventType != XmlPullParser.START_TAG) continue
                when (p.name) {
                    "sldSz" -> {
                        slideCx = p.attrFloat("cx", slideCx).coerceAtLeast(1f)
                        slideCy = p.attrFloat("cy", slideCy).coerceAtLeast(1f)
                    }
                    "sldMasterId" -> if (masterPart == null) {
                        masterPart = opc.targetOf(mainPart, p.relId())
                    }
                    "sldId" -> opc.targetOf(mainPart, p.relId())?.let { slides += Slide(it) }
                    "defaultTextStyle" -> defaultText = parseTxtStyle(p)
                }
            }
        } catch (_: Exception) {
        }
    }

    private fun parseMaster() {
        val mp = masterPart ?: return
        masterShapes = parseShapeTree(mp)
        masterBg = parseBackground(mp)
        // 主题色：母版的 clrMap 把 bg1/tx1 映射到主题 clrScheme 的具体槽位
        parseTheme(mp)
        // 母版的 txStyles 是全局字号底稿（标题 44pt、正文 18pt 等）
        val p = opc.parser(mp) ?: return
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.eventType != XmlPullParser.START_TAG) continue
                when (p.name) {
                    "titleStyle" -> masterTitle = parseTxtStyle(p)
                    "bodyStyle" -> masterBody = parseTxtStyle(p)
                    "otherStyle" -> masterOther = parseTxtStyle(p)
                    "clrMap" -> for (i in 0 until p.attributeCount) {
                        clrMap[p.getAttributeName(i)] = p.getAttributeValue(i) ?: ""
                    }
                }
            }
        } catch (_: Exception) {
        }
    }

    private fun parseTheme(masterPart: String) {
        val themePart = opc.firstByType(masterPart, "theme") ?: return
        val p = opc.parser(themePart) ?: return
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.eventType != XmlPullParser.START_TAG) continue
                if (p.name == "clrScheme") {
                    // 顺序进入 dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink，
                    // 每个槽位里面才是 srgbClr / sysClr
                    val depth = p.depth
                    while (true) {
                        val ev = p.next()
                        if (ev == XmlPullParser.END_DOCUMENT) break
                        if (ev == XmlPullParser.END_TAG && p.depth == depth) break
                        if (ev != XmlPullParser.START_TAG) continue
                        val slot = p.name
                        val color = readSlotColor(p) ?: continue
                        scheme[slot] = color
                    }
                    break
                }
            }
        } catch (_: Exception) {
        }
        if (scheme.isEmpty()) {
            scheme["dk1"] = Color.BLACK; scheme["lt1"] = Color.WHITE
            scheme["dk2"] = 0xFF44546A.toInt(); scheme["lt2"] = 0xFFE7E6E6.toInt()
        }
    }

    /** 主题色槽位（dk1/lt1/accent1…）内部子元素的颜色。 */
    private fun readSlotColor(p: XmlPullParser): Int? {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "srgbClr", "schemeClr", "sysClr", "prstClr" -> return parseColorEl(p)
            }
        }
        return null
    }

    private fun schemeColor(name: String?): Int? {
        if (name == null) return null
        val mapped = clrMap[name] ?: name
        scheme[mapped]?.let { return it }
        return scheme[name]
    }

    // ---------------------------------------------------------------- 形状树

    private fun parseShapeTree(part: String): List<Node> {
        val p = opc.parser(part) ?: return emptyList()
        val out = ArrayList<Node>()
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.at("spTree")) {
                    parseSpTreeChildren(p, out, part)
                    break
                }
            }
        } catch (_: Exception) {
        }
        return out
    }

    private fun parseSpTreeChildren(p: XmlPullParser, out: MutableList<Node>, part: String) {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) return
            if (ev == XmlPullParser.END_TAG && p.depth == depth) return
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "sp" -> out += parseSp(p, part, Kind.SP)
                "pic" -> out += parseSp(p, part, Kind.PIC)
                "cxnSp" -> out += parseSp(p, part, Kind.LINE)
                "graphicFrame" -> out += parseFrame(p, part)
                "grpSp" -> out += parseGroup(p, part)
                else -> {}
            }
        }
    }

    private fun parseGroup(p: XmlPullParser, part: String): Node {
        val node = Node(Kind.GROUP)
        val kids = ArrayList<Node>()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "grpSpPr" -> parseSpPr(p, node)
                "sp" -> kids += parseSp(p, part, Kind.SP)
                "pic" -> kids += parseSp(p, part, Kind.PIC)
                "cxnSp" -> kids += parseSp(p, part, Kind.LINE)
                "graphicFrame" -> kids += parseFrame(p, part)
                "grpSp" -> kids += parseGroup(p, part)
            }
        }
        node.children = kids
        return node
    }

    /** p:sp / p:pic / p:cxnSp 结构相同，只是内容不同，一套解析覆盖三种。 */
    private fun parseSp(p: XmlPullParser, part: String, kind: Kind): Node {
        val node = Node(kind)
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "nvSpPr", "nvPicPr", "nvCxnSpPr" -> parseNv(p, node)
                "spPr" -> parseSpPr(p, node)
                "txBody" -> node.body = parseBody(p)
                "style" -> node.style = parsePhStyle(p)
                "blipFill" -> parseBlipFill(p, node, part)
            }
        }
        return node
    }

    private fun parseFrame(p: XmlPullParser, part: String): Node {
        val node = Node(Kind.FRAME)
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "xfrm" -> node.xfrm = parseXfrm(p)
                "graphic" -> parseGraphic(p, node, part)
                "nvGraphicFramePr" -> parseNv(p, node)
            }
        }
        return node
    }

    private fun parseGraphic(p: XmlPullParser, node: Node, part: String) {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "tbl" -> node.table = parseTable(p)
                "chart" -> node.chartPart = opc.targetOf(part, p.relId())
            }
        }
    }

    private fun parseNv(p: XmlPullParser, node: Node) {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            if (p.name == "ph") {
                node.phType = p.attr("type") ?: "obj"
                node.phIdx = p.attr("idx")?.toIntOrNull()
            }
        }
    }

    private fun parseSpPr(p: XmlPullParser, node: Node) {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "xfrm" -> node.xfrm = parseXfrm(p)
                "prstGeom" -> node.geom = p.attr("prst") ?: "rect"
                "custGeom" -> node.geom = "cust"
                "solidFill", "gradFill", "pattFill", "noFill" -> node.fill = parseFill(p)
                "ln" -> node.line = parseLine(p)
            }
        }
    }

    private fun parseXfrm(p: XmlPullParser): Xfrm {
        val x = Xfrm()
        val depth = p.depth
        x.rotDeg = p.attrFloat("rot", 0f) / 60000f
        x.flipH = p.attr("flipH") == "1"
        x.flipV = p.attr("flipV") == "1"
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "off" -> { x.x = p.attrFloat("x", 0f); x.y = p.attrFloat("y", 0f) }
                "ext" -> { x.cx = p.attrFloat("cx", 0f); x.cy = p.attrFloat("cy", 0f) }
                "chOff" -> { x.chOffX = p.attrFloat("x", 0f); x.chOffY = p.attrFloat("y", 0f) }
                "chExt" -> { x.chExtX = p.attrFloat("cx", 0f); x.chExtY = p.attrFloat("cy", 0f) }
            }
        }
        return x
    }

    private fun parseBlipFill(p: XmlPullParser, node: Node, part: String) {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "blip" -> node.image = opc.targetOf(part, p.relId("embed"))
                "srcRect" -> node.srcRect = RectF(
                    p.attrFloat("l", 0f) / 100000f, p.attrFloat("t", 0f) / 100000f,
                    p.attrFloat("r", 0f) / 100000f, p.attrFloat("b", 0f) / 100000f,
                )
            }
        }
    }

    private fun parsePhStyle(p: XmlPullParser): PhStyle {
        var fillRef: Int? = null
        var lineRef: Int? = null
        var fontRef: Int? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "fillRef" -> fillRef = p.attrInt("idx", 0)
                "lnRef" -> lineRef = p.attrInt("idx", 0)
                "fontRef" -> fontRef = p.attrIdxOrScheme()
            }
        }
        return PhStyle(fillRef, lineRef, fontRef)
    }

    private fun XmlPullParser.attrIdxOrScheme(): Int? {
        val idx = attrInt("idx", -1)
        return if (idx >= 0) idx else null
    }

    // ---------------------------------------------------------------- 填充 / 线 / 颜色

    private fun parseFill(p: XmlPullParser): Fill {
        when (p.name) {
            "noFill" -> { p.skip(); return Fill(none = true) }
            "blipFill" -> {
                // 形状用图片填充不常见，且需要 part 上下文才能解析关系，这里放过
                p.skip()
                return Fill(none = true)
            }
            "gradFill" -> {
                // 渐变只取第一个色标：手机端做成"接近"即可，不值得写完整渐变着色器
                return Fill(solid = firstColorDeep(p) ?: Color.WHITE)
            }
        }
        // solidFill / pattFill
        val c = parseFirstColor(p)
        return Fill(solid = c ?: Color.WHITE)
    }

    /** 在容器内任意深度找第一个颜色（渐变要钻进 gsLst/gs 才拿得到）。 */
    private fun firstColorDeep(p: XmlPullParser): Int? {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "srgbClr", "schemeClr", "sysClr", "prstClr" -> return parseColorEl(p)
            }
        }
        return null
    }

    /** 读容器内的第一个颜色子元素（srgbClr / schemeClr / sysClr / prstClr）。 */
    private fun parseFirstColor(p: XmlPullParser): Int? {
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "srgbClr", "schemeClr", "sysClr", "prstClr" -> return parseColorEl(p)
            }
        }
        return null
    }

    /** 当前元素是颜色元素；顺带应用 lumMod/lumOff 这类修饰。 */
    private fun parseColorEl(p: XmlPullParser): Int? {
        var base = when (p.name) {
            "srgbClr" -> p.attr("val")?.let { parseColor(it, Color.BLACK) }
            "schemeClr" -> schemeColor(p.attr("val"))
            "sysClr" -> p.attr("lastClr")?.let { parseColor(it, Color.BLACK) }
            "prstClr" -> p.attr("val")?.let { parseColor(it, Color.BLACK) }
            else -> null
        }
        var lumMod = 1f
        var lumOff = 0f
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "lumMod" -> lumMod = p.attrFloat("val", 100000f) / 100000f
                "lumOff" -> lumOff = p.attrFloat("val", 0f) / 100000f
            }
        }
        return base?.let { shade(it, lumMod, lumOff) }
    }

    private fun parseLine(p: XmlPullParser): Line {
        val ln = Line()
        ln.widthEmu = p.attrFloat("w", 12700f)
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "noFill" -> { ln.none = true; p.skip() }
                "solidFill", "gradFill" -> parseFirstColor(p)?.let { ln.color = it }
            }
        }
        return ln
    }

    // ---------------------------------------------------------------- 文本

    private fun parseBody(p: XmlPullParser): Body {
        var anchor: String? = null
        var lIns = 91440f; var tIns = 45720f; var rIns = 91440f; var bIns = 45720f
        var wrap = true
        var lst: TxtStyle? = null
        val paras = ArrayList<Para>()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "bodyPr" -> {
                    anchor = p.attr("anchor")
                    lIns = p.attrFloat("lIns", lIns)
                    tIns = p.attrFloat("tIns", tIns)
                    rIns = p.attrFloat("rIns", rIns)
                    bIns = p.attrFloat("bIns", bIns)
                    wrap = p.attr("wrap") != "none"
                    p.skip()
                }
                "lstStyle" -> lst = parseTxtStyle(p)
                "p" -> paras += parsePara(p)
            }
        }
        return Body(paras, anchor, lIns, tIns, rIns, bIns, wrap, lst)
    }

    private fun parseTxtStyle(p: XmlPullParser): TxtStyle {
        val st = TxtStyle()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            val lvl = levelOf(p.name)
            if (lvl > 0) st.lvls[lvl - 1] = parseLvl(p)
        }
        return st
    }

    private fun levelOf(tag: String): Int {
        if (!tag.startsWith("lvl") || !tag.endsWith("pPr")) return 0
        return tag.substring(3, tag.length - 3).toIntOrNull() ?: 0
    }

    private fun parseLvl(p: XmlPullParser): LvlStyle {
        var sz: Int? = null; var bold: Boolean? = null; var italic: Boolean? = null
        var color: Int? = null; var bullet: String? = null; var hasBuNone = false
        val algn = p.attr("algn")
        val marL = p.attr("marL")?.toFloatOrNull()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "buNone" -> { hasBuNone = true; p.skip() }
                "buChar" -> bullet = p.attr("char") ?: "•"
                "defRPr" -> {
                    val d = parseRPr(p)
                    sz = d.sz ?: sz; bold = d.bold ?: bold; italic = d.italic ?: italic; color = d.color ?: color
                }
            }
        }
        return LvlStyle(sz, bold, italic, color, algn, marL, if (hasBuNone) "" else bullet)
    }

    private class RunProps(val sz: Int?, val bold: Boolean?, val italic: Boolean?, val underline: Boolean?, val color: Int?)

    private fun parseRPr(p: XmlPullParser): RunProps {
        var sz = p.attr("sz")?.toIntOrNull()
        var bold = p.attr("b")?.let { it == "1" || it == "true" }
        var italic = p.attr("i")?.let { it == "1" || it == "true" }
        val underline = p.attr("u")?.let { it != "none" }
        var color: Int? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "solidFill" -> color = parseFirstColor(p) ?: color
                "latin", "ea", "cs" -> p.skip()
            }
        }
        return RunProps(sz, bold, italic, underline, color)
    }

    private fun parsePara(p: XmlPullParser): Para {
        var algn: String? = null; var marL: Float? = null; var indent: Float? = null
        var lvl = 0; var bullet: String? = null; var linePct: Float? = null; var spaceBef: Float? = null
        val runs = ArrayList<Run>()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "pPr" -> {
                    algn = p.attr("algn")
                    marL = p.attr("marL")?.toFloatOrNull()
                    indent = p.attr("indent")?.toFloatOrNull()
                    lvl = (p.attr("lvl")?.toIntOrNull() ?: 0) + 1
                    val d = p.depth
                    while (true) {
                        val e2 = p.next()
                        if (e2 == XmlPullParser.END_DOCUMENT) break
                        if (e2 == XmlPullParser.END_TAG && p.depth == d) break
                        if (e2 != XmlPullParser.START_TAG) continue
                        when (p.name) {
                            "lnSpc" -> linePct = parseSpcPct(p)
                            "spcBef" -> spaceBef = parseSpcPts(p)
                            "buChar" -> bullet = p.attr("char")
                            "buNone" -> { bullet = ""; p.skip() }
                        }
                    }
                }
                "r" -> parseRun(p)?.let { runs += it }
                "fld" -> parseRun(p)?.let { runs += it }
                "br" -> runs += Run("\n")
            }
        }
        return Para(runs, algn, marL, indent, lvl, bullet, linePct, spaceBef)
    }

    private fun parseSpcPct(p: XmlPullParser): Float? {
        var v: Float? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            if (p.name == "spcPct") v = p.attrFloat("val", 100000f) / 1000f
        }
        return v
    }

    private fun parseSpcPts(p: XmlPullParser): Float? {
        var v: Float? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            if (p.name == "spcPts") v = p.attrFloat("val", 0f) / 100f
        }
        return v
    }

    private fun parseRun(p: XmlPullParser): Run? {
        var props = RunProps(null, null, null, null, null)
        val sb = StringBuilder()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "rPr" -> props = parseRPr(p)
                "t" -> sb.append(p.plainText())
            }
        }
        if (sb.isEmpty()) return null
        return Run(sb.toString(), props.sz, props.bold, props.italic, props.underline, props.color)
    }

    // ---------------------------------------------------------------- 表格

    private fun parseTable(p: XmlPullParser): Table {
        val colW = ArrayList<Float>()
        val rows = ArrayList<List<Cell>>()
        val rowH = ArrayList<Float>()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "gridCol" -> colW += p.attrFloat("w", 0f)
                "tr" -> rows += parseRow(p, rowH)
            }
        }
        if (colW.isEmpty()) repeat(3) { colW += 1f }
        return Table(colW.toFloatArray(), rows, rowH.toFloatArray())
    }

    private fun parseRow(p: XmlPullParser, rowH: MutableList<Float>): List<Cell> {
        rowH += p.attrFloat("h", 370840f)
        val cells = ArrayList<Cell>()
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            if (p.name == "tc") cells += parseCell(p)
        }
        return cells
    }

    private fun parseCell(p: XmlPullParser): Cell {
        var body: Body? = null
        var fill: Int? = null
        var span = 1
        var rspan = 1
        var anchor: String? = null
        val depth = p.depth
        while (true) {
            val ev = p.next()
            if (ev == XmlPullParser.END_DOCUMENT) break
            if (ev == XmlPullParser.END_TAG && p.depth == depth) break
            if (ev != XmlPullParser.START_TAG) continue
            when (p.name) {
                "txBody" -> body = parseBody(p)
                "tcPr" -> {
                    anchor = p.attr("anchor")
                    span = p.attrInt("gridSpan", 1)
                    rspan = p.attrInt("rowSpan", 1)
                    val d = p.depth
                    while (true) {
                        val e2 = p.next()
                        if (e2 == XmlPullParser.END_DOCUMENT) break
                        if (e2 == XmlPullParser.END_TAG && p.depth == d) break
                        if (e2 != XmlPullParser.START_TAG) continue
                        if (p.name == "solidFill") fill = parseFirstColor(p)
                    }
                }
            }
        }
        return Cell(body, fill, span, rspan, anchor)
    }

    // ---------------------------------------------------------------- 背景

    private fun parseBackground(part: String): Fill? {
        val p = opc.parser(part) ?: return null
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.at("bg")) {
                    val d = p.depth
                    while (true) {
                        val ev = p.next()
                        if (ev == XmlPullParser.END_DOCUMENT) break
                        if (ev == XmlPullParser.END_TAG && p.depth == d) break
                        if (ev != XmlPullParser.START_TAG) continue
                        when (p.name) {
                            "bgPr" -> {
                                val d2 = p.depth
                                while (true) {
                                    val e2 = p.next()
                                    if (e2 == XmlPullParser.END_DOCUMENT) break
                                    if (e2 == XmlPullParser.END_TAG && p.depth == d2) break
                                    if (e2 != XmlPullParser.START_TAG) continue
                                    if (p.name.endsWith("Fill")) return parseFill(p)
                                }
                            }
                            "bgRef" -> {
                                // 引用主题 bgFillStyleLst 的第 idx-1001 项
                                val idx = p.attrInt("idx", 1001) - 1001
                                parseFirstColor(p)?.let { return Fill(solid = it) }
                                themeBgFill(idx)?.let { return it }
                                return null
                            }
                        }
                    }
                    return null
                }
            }
        } catch (_: Exception) {
        }
        return null
    }

    /** 主题里的背景填充列表（fmtScheme/bgFillStyleLst）。 */
    private fun themeBgFill(index: Int): Fill? {
        val mp = masterPart ?: return null
        val themePart = opc.firstByType(mp, "theme") ?: return null
        val p = opc.parser(themePart) ?: return null
        try {
            while (p.next() != XmlPullParser.END_DOCUMENT) {
                if (p.at("bgFillStyleLst")) {
                    var i = 0
                    val d = p.depth
                    while (true) {
                        val ev = p.next()
                        if (ev == XmlPullParser.END_DOCUMENT) break
                        if (ev == XmlPullParser.END_TAG && p.depth == d) break
                        if (ev != XmlPullParser.START_TAG) continue
                        if (!p.name.endsWith("Fill")) continue
                        if (i == index) return parseFill(p)
                        p.skip()
                        i++
                    }
                }
            }
        } catch (_: Exception) {
        }
        return null
    }

    // ---------------------------------------------------------------- 渲染

    override fun renderPage(index: Int, targetWidth: Int): Bitmap =
        render(index, targetWidth.coerceAtLeast(32).toFloat())

    override fun renderPageDpi(index: Int, dpi: Int): Bitmap {
        // 幻灯片物理宽度（英寸）= cx / 914400；像素 = 英寸 × DPI
        val w = (slideCx / 914400f) * dpi
        var px = max(32f, w)
        val h = px * slideCy / slideCx
        // 600 DPI 的大幅面会爆内存：整体限到 4000 万像素
        val area = px * h
        if (area > MAX_PIXELS) {
            val f = kotlin.math.sqrt(MAX_PIXELS / area)
            px *= f
        }
        return render(index, px)
    }

    private fun render(index: Int, widthPx: Float): Bitmap = synchronized(lock) {
        val slide = slides.getOrNull(index) ?: throw IllegalStateException("页码越界")
        val shapes = slide.shapes ?: parseShapeTree(slide.part).also { slide.shapes = it }
        val layoutPart = layoutOf(slide)
        val layoutShapes = layoutPart?.let { layoutShapesCache.getOrPut(it) { parseShapeTree(it) } } ?: emptyList()

        val w = widthPx.toInt().coerceAtLeast(32)
        val h = max(16, (w * slideCy / slideCx).toInt())
        pxPerEmu = w / slideCx
        val bmp = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888)
        val c = Canvas(bmp)
        c.scale(w / slideCx, h / slideCy)   // 之后一律用 EMU 坐标

        // 背景：页 → 版式 → 母版
        val bg = slide.bg
            ?: layoutPart?.let { layoutBgCache.getOrPut(it) { parseBackground(it) } }
            ?: masterBg
        drawFillRect(c, RectF(0f, 0f, slideCx, slideCy), bg)

        // 母版/版式的装饰形状先画（占位符一律跳过，里面只有"单击此处编辑…"）
        val lIdx = layoutPart?.let { layoutIdxCache.getOrPut(it) { indexOf(layoutShapes) } } ?: emptyIdx
        drawNodes(c, masterShapes, emptyIdx, emptyIdx, isMaster = true)
        drawNodes(c, layoutShapes, emptyIdx, masterIdx, isMaster = true)
        drawNodes(c, shapes, lIdx, masterIdx, isMaster = false)

        bmp
    }

    private fun layoutOf(slid: Slide): String? {
        if (!slid.layoutLoaded) {
            slid.layoutLoaded = true
            slid.layout = opc.firstByType(slid.part, "slideLayout")
        }
        return slid.layout
    }

    /** 版式/母版里的占位符索引：先按 idx 精确匹配，再按 type 兜底。 */
    private class PhIndex(val byIdx: Map<Int, Node>, val byType: Map<String, Node>)

    private val emptyIdx = PhIndex(emptyMap(), emptyMap())
    private val layoutIdxCache = HashMap<String, PhIndex>()
    private val masterIdx: PhIndex by lazy { indexOf(masterShapes) }

    private fun indexOf(nodes: List<Node>): PhIndex {
        val byIdx = HashMap<Int, Node>()
        val byType = HashMap<String, Node>()
        fun walk(list: List<Node>) {
            for (n in list) {
                if (n.phType != null) {
                    n.phIdx?.let { byIdx.putIfAbsent(it, n) }
                    byType.putIfAbsent(n.phType!!, n)
                }
                n.children?.let { walk(it) }
            }
        }
        walk(nodes)
        return PhIndex(byIdx, byType)
    }

    private fun lookupPh(node: Node, idx: PhIndex): Node? {
        node.phIdx?.let { i -> idx.byIdx[i]?.let { return it } }
        node.phType?.let { t -> idx.byType[t]?.let { return it } }
        // "obj"/"body" 这类泛型占位符常见 idx 对不上但顺序一致，放宽到第一个 body
        if (node.phType != null && (node.phType == "obj" || node.phType == "body")) {
            return idx.byType["body"] ?: idx.byType["obj"]
        }
        return null
    }

    private fun drawNodes(c: Canvas, nodes: List<Node>, lIdx: PhIndex, mIdx: PhIndex, isMaster: Boolean) {
        for (n in nodes) {
            try {
                drawNode(c, n, lIdx, mIdx, isMaster)
            } catch (e: Throwable) {
                // 单个形状画挂了不影响整页
                android.util.Log.w(TAG, "draw fail ${n.kind}/${n.phType}", e)
            }
        }
    }

    private fun drawNode(c: Canvas, node: Node, lIdx: PhIndex, mIdx: PhIndex, isMaster: Boolean) {
        // 母版/版式里的占位符一律不画：它们只是"单击此处编辑…"的提示
        if (isMaster && node.phType != null) return

        // 继承链：本形状 → 版式同位置占位符 → 母版同位置占位符。
        // 三层都可能只有空 spPr，坐标要一路往下找（样例正是这种情况）。
        val chain = ArrayList<Node>(3)
        chain += node
        if (node.phType != null) {
            lookupPh(node, lIdx)?.let { if (chain.none { o -> o === it }) chain += it }
            lookupPh(node, mIdx)?.let { if (chain.none { o -> o === it }) chain += it }
        }
        val xf = chain.firstNotNullOfOrNull { it.xfrm } ?: return
        if (xf.cx <= 0f || xf.cy <= 0f) return
        val rect = RectF(xf.x, xf.y, xf.x + xf.cx, xf.y + xf.cy)

        c.save()
        if (xf.rotDeg != 0f) c.rotate(xf.rotDeg, rect.centerX(), rect.centerY())

        when (node.kind) {
            Kind.GROUP -> {
                val kids = node.children ?: return
                c.save()
                // 子形状坐标系 chOff/chExt → 组坐标 off/ext
                val sx = if (xf.chExtX > 0f) xf.cx / xf.chExtX else 1f
                val sy = if (xf.chExtY > 0f) xf.cy / xf.chExtY else 1f
                c.translate(xf.x, xf.y)
                c.scale(sx, sy)
                c.translate(-xf.chOffX, -xf.chOffY)
                drawNodes(c, kids, lIdx, mIdx, isMaster)
                c.restore()
            }
            Kind.PIC -> {
                val part = node.image
                val bmp = part?.let { imageOf(it) }
                if (bmp != null) drawPicture(c, bmp, rect, node.srcRect)
                else drawPlaceholder(c, rect, "［图片］")
            }
            Kind.TABLE -> {
                val t = node.table
                if (t != null) drawTable(c, rect, t)
            }
            Kind.FRAME -> {
                val t = node.table
                if (t != null) drawTable(c, rect, t)
                else drawPlaceholder(c, rect, "［图表］")
            }
            Kind.LINE -> {
                val ln = node.line ?: Line(color = 0xFF808080.toInt(), widthEmu = 12700f)
                if (!ln.none) {
                    val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                        color = ln.color
                        strokeWidth = max(ln.widthEmu, 6350f)
                        style = Paint.Style.STROKE
                    }
                    c.drawLine(rect.left, rect.top, rect.right, rect.bottom, paint)
                }
            }
            Kind.SP -> {
                val fill = node.fill ?: chain.firstNotNullOfOrNull { it.fill } ?: fillFromStyle(node.style)
                val path = geomPath(node.geom, rect)
                if (fill != null && !fill.none) drawFillPath(c, path, rect, fill)
                val ln = node.line ?: chain.firstNotNullOfOrNull { it.line } ?: lineFromStyle(node.style)
                if (ln != null && !ln.none && ln.widthEmu > 0f) {
                    val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                        color = ln.color
                        strokeWidth = ln.widthEmu
                        style = Paint.Style.STROKE
                    }
                    c.drawPath(path, paint)
                }
                drawText(c, rect, node.body, node, chain)
            }
        }
        c.restore()
    }

    private fun drawPlaceholder(c: Canvas, rect: RectF, text: String) {
        val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            color = 0xFFEDEFF3.toInt()
            style = Paint.Style.FILL
        }
        c.drawRoundRect(rect, 60960f, 60960f, paint)
        val tp = TextPaint(Paint.ANTI_ALIAS_FLAG).apply {
            color = 0xFF98A2B3.toInt()
            textSize = 228600f
            textAlign = Paint.Align.CENTER
        }
        c.drawText(text, rect.centerX(), rect.centerY() + 80000f, tp)
    }

    private fun drawPicture(c: Canvas, bmp: Bitmap, rect: RectF, src: RectF?) {
        val paint = Paint(Paint.FILTER_BITMAP_FLAG or Paint.ANTI_ALIAS_FLAG)
        if (src == null || (src.left == 0f && src.top == 0f && src.right == 0f && src.bottom == 0f)) {
            c.drawBitmap(bmp, null, rect, paint)
        } else {
            val r = Rect(
                (bmp.width * src.left).toInt().coerceIn(0, bmp.width - 1),
                (bmp.height * src.top).toInt().coerceIn(0, bmp.height - 1),
                (bmp.width * (1f - src.right)).toInt().coerceIn(1, bmp.width),
                (bmp.height * (1f - src.bottom)).toInt().coerceIn(1, bmp.height),
            )
            if (r.width() > 0 && r.height() > 0) c.drawBitmap(bmp, r, rect, paint)
        }
    }

    private fun imageOf(part: String): Bitmap? {
        imgCache.get(part)?.let { return it }
        val bytes = opc.bytes(part) ?: return null
        val bmp = try {
            BitmapFactory.decodeByteArray(bytes, 0, bytes.size)
        } catch (_: Throwable) {
            null
        } ?: return null
        imgCache.put(part, bmp)
        return bmp
    }

    private fun fillFromStyle(s: PhStyle?): Fill? {
        val idx = s?.fillRef ?: return null
        if (idx <= 0) return null
        val accent = schemeColor("accent${((idx - 1) % 6) + 1}") ?: return null
        return Fill(solid = accent)
    }

    private fun lineFromStyle(s: PhStyle?): Line? {
        val idx = s?.lineRef ?: return null
        if (idx <= 0) return null
        val accent = schemeColor("accent${((idx - 1) % 6) + 1}") ?: return null
        return Line(color = accent, widthEmu = 12700f)
    }

    private fun geomPath(geom: String, rect: RectF): Path {
        val path = Path()
        when (geom) {
            "ellipse" -> path.addOval(rect, Path.Direction.CW)
            "roundRect" -> path.addRoundRect(rect, min(rect.width(), rect.height()) * 0.1f, min(rect.width(), rect.height()) * 0.1f, Path.Direction.CW)
            "triangle" -> {
                path.moveTo(rect.centerX(), rect.top)
                path.lineTo(rect.right, rect.bottom)
                path.lineTo(rect.left, rect.bottom)
                path.close()
            }
            "line" -> {
                path.moveTo(rect.left, rect.top)
                path.lineTo(rect.right, rect.bottom)
            }
            else -> path.addRect(rect, Path.Direction.CW)
        }
        return path
    }

    private fun drawFillRect(c: Canvas, rect: RectF, fill: Fill?) {
        val color = fill?.solid ?: Color.WHITE
        c.drawRect(rect, Paint().apply { this.color = color })
    }

    private fun drawFillPath(c: Canvas, path: Path, rect: RectF, fill: Fill) {
        if (fill.imagePart != null) {
            val bmp = imageOf(fill.imagePart) ?: return
            c.save()
            c.clipPath(path)
            c.drawBitmap(bmp, null, rect, Paint(Paint.FILTER_BITMAP_FLAG))
            c.restore()
            return
        }
        c.drawPath(path, Paint(Paint.ANTI_ALIAS_FLAG).apply {
            color = fill.solid ?: Color.WHITE
            style = Paint.Style.FILL
        })
    }

    // ---------------------------------------------------------------- 表格绘制

    private fun drawTable(c: Canvas, rect: RectF, t: Table) {
        val totalW = t.colW.sum().takeIf { it > 0f } ?: rect.width()
        val sx = rect.width() / totalW
        val rowHeights = FloatArray(t.rows.size) { i ->
            t.rowH.getOrNull(i) ?: 370840f
        }
        val totalH = rowHeights.sum().takeIf { it > 0f } ?: rect.height()
        val sy = rect.height() / totalH

        val gridPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            color = 0xFFBFBFBF.toInt()
            style = Paint.Style.STROKE
            strokeWidth = 12700f
        }
        var y = rect.top
        t.rows.forEachIndexed { ri, row ->
            var x = rect.left
            val rh = rowHeights[ri] * sy
            row.forEachIndexed { ci, cell ->
                var cw = 0f
                for (k in 0 until cell.colSpan) cw += (t.colW.getOrNull(ci + k) ?: 0f) * sx
                val cr = RectF(x, y, x + cw, y + rh)
                c.drawRect(cr, Paint().apply { color = cell.fill ?: Color.WHITE })
                c.drawRect(cr, gridPaint)
                cell.text?.let { body -> drawText(c, cr, body, null, emptyList(), pad = 27432f) }
                x += cw
            }
            y += rh
        }
        c.drawRect(rect, gridPaint)
    }

    // ---------------------------------------------------------------- 文字排版

    /** 逐级向下找：前者有值用前者，没有就落到下一层。 */
    private fun mergeLvl(a: LvlStyle?, b: LvlStyle?): LvlStyle? = when {
        a == null -> b
        b == null -> a
        else -> a.merge(b)
    }

    private fun drawText(
        c: Canvas, box: RectF, body: Body?, node: Node?,
        chain: List<Node>, pad: Float? = null,
    ) {
        if (body == null || body.paras.isEmpty()) return
        // 统一换算到像素：EMU 里字号动辄几十万，StaticLayout/Skia 在那个量级不出字，
        // 而 1pt = 12700 EMU 本来就只是一次乘法。
        val s = pxPerEmu
        if (s <= 0f) return
        val l = (box.left + (pad ?: body.lIns)) * s
        val t = (box.top + (pad ?: body.tIns)) * s
        val r = (box.right - (pad ?: body.rIns)) * s
        val b = (box.bottom - (pad ?: body.bIns)) * s
        val availW = r - l
        if (availW <= 1f || b - t <= 1f) return

        // 字号/颜色的继承链：本形状 lstStyle → 版式占位符 → 母版占位符 → 母版 txStyles
        val styleOfTitle = node?.phType == "title" || node?.phType == "ctrTitle"
        val styleOfBody = node?.phType == "body" || node?.phType == "subTitle" || node?.phType == "obj"
        fun masterStyle(): TxtStyle? =
            if (styleOfTitle) masterTitle else if (styleOfBody) masterBody else masterOther

        val inherited = chain.drop(1)

        fun lvlStyle(lvl: Int): LvlStyle? {
            var acc: LvlStyle? = body.lstStyle?.get(lvl)
            for (anc in inherited) acc = mergeLvl(acc, anc.body?.lstStyle?.get(lvl))
            acc = mergeLvl(acc, masterStyle()?.get(lvl))
            acc = mergeLvl(acc, defaultText?.get(lvl))
            return acc
        }

        // 先排一遍量高度，超出就整体缩号（PPT 的"溢出缩小"）
        var autofit = 1f
        repeat(6) {
            val laid = layoutParas(body, availW, { lvl -> lvlStyle(lvl) }, autofit, s)
            val total = laid.sumOf { (it.layout.height + it.gap).toDouble() }.toFloat()
            if (total <= b - t || autofit <= 0.55f) {
                var y = t
                when (body.anchor) {
                    "ctr" -> y = t + ((b - t) - total) / 2f
                    "b" -> y = b - total
                }
                c.save()
                c.scale(1f / s, 1f / s)      // 抵消 render() 的整体缩放，之后按像素画
                for (p in laid) {
                    y += p.gap
                    p.bullet?.let { bl ->
                        c.save()
                        c.translate(l + p.bdx, y)
                        bl.draw(c)
                        c.restore()
                    }
                    c.save()
                    c.translate(l + p.dx, y)
                    p.layout.draw(c)
                    c.restore()
                    y += p.layout.height
                }
                c.restore()
                return
            }
            autofit *= 0.88f
        }
    }

    /**
     * 一段排好的文字：正文布局 + 段前间距 + 左缩进（像素）。
     * bullet 是单独排的项目符号布局（正文各行对齐 marL，符号挂在 marL+indent 处）。
     */
    private class LaidPara(
        val layout: StaticLayout,
        val gap: Float,
        val dx: Float,
        val bullet: StaticLayout? = null,
        val bdx: Float = 0f,
    )

    /**
     * 返回「排好的段落」+ 该段之前的间距（像素）。
     * availW 与返回的高度都是像素，autofit 是溢出缩小系数。
     */
    private fun layoutParas(
        body: Body, availW: Float,
        lvlStyle: (Int) -> LvlStyle?,
        autofit: Float,
        pxPerEmu: Float,
    ): List<LaidPara> {
        val out = ArrayList<LaidPara>()
        // 百分之一磅 → 像素：sz/100 * 72pt/英寸? 不用，1pt = 12700 EMU，乘缩放即可
        fun px(sz: Int) = sz * 127f * pxPerEmu * autofit

        body.paras.forEachIndexed { pi, para ->
            val lvl = para.lvl.coerceIn(1, 9)
            val def = lvlStyle(lvl)
            val defSz = px(def?.sz ?: 1800)
            val paraColor = def?.color ?: Color.BLACK
            val defBold = def?.bold ?: false
            val bulletChar = para.bullet ?: def?.bullet
            val bulletGlyph: String? = bulletChar?.takeIf { it.isNotEmpty() }
            val showBullet = bulletGlyph != null
            // marL 是段落左缩进，indent 是首行相对缩进（负数 = 悬挂，项目符号挂在外侧）
            val marL = (para.marL ?: def?.marL ?: 0f) * pxPerEmu
            val indent = (para.indent ?: 0f) * pxPerEmu

            // 项目符号不再塞进正文串里：那样换行后的文字会顶着符号左对齐，和 PPT 不一样
            val sp = SpannableString(buildString {
                para.runs.forEach { append(it.text) }
            })
            var pos = 0
            for (run in para.runs) {
                val start = pos
                val end = pos + run.text.length
                pos = end
                if (end <= start) continue
                val sz = if (run.sz != null) px(run.sz) else defSz
                sp.setSpan(AbsoluteSizeSpan(sz.toInt().coerceAtLeast(1)), start, end, 0)
                sp.setSpan(ForegroundColorSpan(run.color ?: paraColor), start, end, 0)
                val bold = run.bold ?: defBold
                val italic = run.italic ?: false
                if (bold || italic) {
                    val st = when {
                        bold && italic -> Typeface.BOLD_ITALIC
                        bold -> Typeface.BOLD
                        else -> Typeface.ITALIC
                    }
                    sp.setSpan(StyleSpan(st), start, end, 0)
                }
                if (run.underline == true) sp.setSpan(UnderlineSpan(), start, end, 0)
            }

            val paint = TextPaint(Paint.ANTI_ALIAS_FLAG).apply {
                textSize = defSz
                color = paraColor
                isAntiAlias = true
            }
            val align = when (para.algn ?: def?.algn) {
                "ctr" -> Layout.Alignment.ALIGN_CENTER
                "r" -> Layout.Alignment.ALIGN_OPPOSITE
                else -> Layout.Alignment.ALIGN_NORMAL
            }
            // 悬挂缩进：正文各行统一从 marL 起排，项目符号单独排在 marL+indent
            var textDx = marL.coerceIn(0f, availW * 0.6f)
            var bulletDx = (marL + indent).coerceIn(0f, availW * 0.6f)
            if (showBullet && textDx <= bulletDx + 1f) {
                // 作者没给负缩进（或压根没写缩进）：给符号让出一段空间，避免压到正文
                textDx = (paint.measureText(bulletGlyph ?: "") + defSz * 0.32f)
                    .coerceIn(0f, availW * 0.6f)
                bulletDx = 0f
            }
            val width = (availW - textDx).coerceAtLeast(1f).toInt()
            val layout = StaticLayout.Builder
                .obtain(sp, 0, sp.length, paint, width)
                .setAlignment(align)
                .setLineSpacing(0f, (para.linePct ?: 100f) / 100f)
                .setIncludePad(false)
                .setBreakStrategy(Layout.BREAK_STRATEGY_HIGH_QUALITY)
                .build()
            // 符号用同一支画笔排，行高/基线自然和正文首行对齐
            val bulletLayout = if (bulletGlyph != null) {
                val bsp = SpannableString(bulletGlyph)
                StaticLayout.Builder
                    .obtain(bsp, 0, bsp.length, paint,
                        (availW - bulletDx).coerceAtLeast(1f).toInt())
                    .setAlignment(Layout.Alignment.ALIGN_NORMAL)
                    .setLineSpacing(0f, (para.linePct ?: 100f) / 100f)
                    .setIncludePad(false)
                    .build()
            } else null
            // 段前间距：显式 spcBef 优先，否则除首段外留 0.18 行呼吸感
            val before = (para.spaceBeforePt ?: 0f) * 12700f * pxPerEmu * autofit +
                (if (pi > 0) defSz * 0.18f else 0f)
            out += LaidPara(layout, before, textDx, bulletLayout, bulletDx)
        }
        return out
    }

    override fun close() {
        imgCache.evictAll()
        opc.close()
    }

    companion object {
        private const val MAX_PIXELS = 40_000_000f
        private const val TAG = "MuFileTools"
    }
}
