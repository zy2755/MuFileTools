package com.mutools.filetools.core

import android.content.Context
import android.net.Uri
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory
import java.io.Closeable
import java.io.File
import java.io.InputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipFile

/**
 * OPC（Open Packaging Conventions）容器：docx / pptx / xlsx 共用的外壳。
 *
 * 三种格式都是「zip + [Content_Types].xml + 一堆 .rels」：
 *   - zip 里每个文件叫一个 **part**（如 `ppt/slides/slide1.xml`）
 *   - 每个 part 旁边有 `_rels/<文件名>.rels`，登记「关系 Id → 目标 part」
 *   - 主部件（document/presentation/workbook）由根 `_rels/.rels` 的
 *     `officeDocument` 关系指出，**不能靠猜文件名**
 *
 * SAF 只给 InputStream，而 zip 需要随机读（切页、反复取图片），
 * 所以先整份复制到 cacheDir 再 `ZipFile` 打开；close() 时删掉副本。
 */
class Opc private constructor(private val tmp: File) : Closeable {

    private val zip = ZipFile(tmp)
    private val relsCache = HashMap<String, Map<String, Rel>>()

    class Rel(val id: String, val type: String, val target: String, val external: Boolean)

    // ---------------------------------------------------------------- 读取

    fun has(part: String): Boolean = zip.getEntry(part)?.let { !it.isDirectory } == true

    fun entry(part: String): ZipEntry? = zip.getEntry(part)

    fun stream(part: String): InputStream? {
        val e = zip.getEntry(part) ?: return null
        return zip.getInputStream(e)
    }

    fun bytes(part: String): ByteArray? = stream(part)?.use { it.readBytes() }

    /** part 文本。XML 都是 UTF-8；真正的编码声明交给解析器，这里只去掉 BOM。 */
    fun text(part: String): String? {
        val b = bytes(part) ?: return null
        val off = if (b.size >= 3 && b[0] == 0xEF.toByte() && b[1] == 0xBB.toByte() && b[2] == 0xBF.toByte()) 3 else 0
        return String(b, off, b.size - off, Charsets.UTF_8)
    }

    fun parser(part: String): XmlPullParser? {
        val s = stream(part) ?: return null
        val p = XmlPullParserFactory.newInstance().apply { isNamespaceAware = true }.newPullParser()
        p.setInput(s, "UTF-8")
        return p
    }

    // ---------------------------------------------------------------- 关系

    /** part 对应的 .rels 路径：`ppt/a/b.xml` → `ppt/a/_rels/b.xml.rels`；根部件 → `_rels/.rels`。 */
    fun relsPathOf(part: String): String {
        if (part.isEmpty()) return "_rels/.rels"
        val dir = part.substringBeforeLast('/', "")
        val base = part.substringAfterLast('/')
        return if (dir.isEmpty()) "_rels/$base.rels" else "$dir/_rels/$base.rels"
    }

    fun rels(part: String): Map<String, Rel> = synchronized(relsCache) {
        relsCache.getOrPut(part) { parseRels(part) }
    }

    private fun parseRels(part: String): Map<String, Rel> {
        val p = parser(relsPathOf(part)) ?: return emptyMap()
        val out = HashMap<String, Rel>()
        try {
            var ev = p.eventType
            while (ev != XmlPullParser.END_DOCUMENT) {
                if (ev == XmlPullParser.START_TAG && p.name == "Relationship") {
                    val id = p.attr("Id") ?: ""
                    val target = p.attr("Target") ?: ""
                    val type = p.attr("Type") ?: ""
                    val ext = p.attr("TargetMode")?.equals("External", true) == true
                    if (id.isNotEmpty() && target.isNotEmpty()) {
                        out[id] = Rel(id, type, if (ext) target else resolve(part, target), ext)
                    }
                }
                ev = p.next()
            }
        } catch (_: Exception) {
        }
        return out
    }

    /** 关系 Id → 解析后的 part 名（外部链接返回 null）。 */
    fun targetOf(part: String, relId: String?): String? {
        if (relId == null) return null
        val r = rels(part)[relId] ?: return null
        return if (r.external) null else r.target
    }

    /** 按关系类型找第一个目标（一个 part 可能有多个同类型关系）。 */
    fun firstByType(part: String, typeSuffix: String): String? =
        rels(part).values.firstOrNull { it.type.endsWith(typeSuffix) && !it.external }?.target

    /** 相对路径解析：`ppt/presentation.xml` + `slides/slide1.xml` → `ppt/slides/slide1.xml`。 */
    private fun resolve(basePart: String, target: String): String {
        if (target.startsWith("/")) return target.substring(1)
        val stack = ArrayList<String>()
        val dir = basePart.substringBeforeLast('/', "")
        (if (dir.isEmpty()) "" else "$dir/").split('/').forEach { if (it.isNotEmpty()) stack += it }
        target.split('/').forEach { seg ->
            when (seg) {
                "", "." -> {}
                ".." -> if (stack.isNotEmpty()) stack.removeAt(stack.size - 1)
                else -> stack += seg
            }
        }
        return stack.joinToString("/")
    }

    override fun close() {
        runCatching { zip.close() }
        runCatching { tmp.delete() }
    }

    companion object {
        /** 把 SAF 的 uri 复制到缓存目录再打开；文件名带时间戳避免多实例互踩。 */
        fun open(context: Context, uri: Uri, suffix: String): Opc {
            val tmp = File.createTempFile("opc-", suffix, context.cacheDir)
            val input = context.contentResolver.openInputStream(uri)
                ?: throw IllegalStateException("无法读取文件")
            input.use { src -> tmp.outputStream().use { src.copyTo(it, 64 * 1024) } }
            return Opc(tmp)
        }
    }
}

// ---------------------------------------------------------------- XML 小工具

/** 取不带命名空间的属性（XmlPullParser 开了 namespace 后 name 是 local-name）。 */
fun XmlPullParser.attr(name: String): String? = getAttributeValue(null, name)

fun XmlPullParser.attrInt(name: String, def: Int): Int = attr(name)?.trim()?.toIntOrNull() ?: def

fun XmlPullParser.attrFloat(name: String, def: Float): Float = attr(name)?.trim()?.toFloatOrNull() ?: def

fun XmlPullParser.attrBool(name: String, def: Boolean = false): Boolean =
    attr(name)?.let { it == "1" || it.equals("true", true) } ?: def

/** `r:id` / `r:embed` 这类关系引用：前缀可能是 r，也可能被重命名，按「命名空间含 relationships」认。 */
fun XmlPullParser.relId(name: String = "id"): String? {
    for (i in 0 until attributeCount) {
        if (getAttributeName(i) == name) {
            val ns = getAttributeNamespace(i) ?: continue
            if (ns.contains("relationships")) return getAttributeValue(i)
        }
    }
    return null
}

fun XmlPullParser.at(tag: String): Boolean = eventType == XmlPullParser.START_TAG && name == tag

fun XmlPullParser.atEnd(tag: String): Boolean = eventType == XmlPullParser.END_TAG && name == tag

/** 跳过当前元素（含所有子元素），停在它的 END_TAG 之后。 */
fun XmlPullParser.skip(): XmlPullParser {
    var depth = 1
    while (depth > 0) {
        when (next()) {
            XmlPullParser.START_TAG -> depth++
            XmlPullParser.END_TAG -> depth--
            XmlPullParser.END_DOCUMENT -> return this
        }
    }
    return this
}

/** 读当前元素的纯文本（遇 END_TAG 停）。 */
fun XmlPullParser.plainText(): String {
    val sb = StringBuilder()
    var depth = 1
    while (depth > 0) {
        when (next()) {
            XmlPullParser.START_TAG -> depth++
            XmlPullParser.END_TAG -> depth--
            XmlPullParser.TEXT -> if (depth == 1) sb.append(text)
            XmlPullParser.END_DOCUMENT -> return sb.toString()
        }
    }
    return sb.toString()
}

/** 颜色解析：`FF0000` / `00FFFFFF`（Excel 写 8 位但首字节是保留位，按不透明处理）。 */
fun parseColor(s: String?, def: Int = 0xFF000000.toInt()): Int {
    if (s == null) return def
    val v = s.trim().removePrefix("#")
    return when (v.length) {
        8 -> (0xFF shl 24) or (v.substring(2).toIntOrNull(16) ?: return def)
        6 -> (0xFF shl 24) or (v.toIntOrNull(16) ?: return def)
        3 -> {
            val r = v[0].toString().repeat(2).toIntOrNull(16) ?: return def
            val g = v[1].toString().repeat(2).toIntOrNull(16) ?: return def
            val b = v[2].toString().repeat(2).toIntOrNull(16) ?: return def
            (0xFF shl 24) or (r shl 16) or (g shl 8) or b
        }
        else -> def
    }
}

/** 亮度调节（对应 OOXML 的 lumMod/lumOff），用于主题色的深浅变体。 */
fun shade(color: Int, lumMod: Float = 1f, lumOff: Float = 0f): Int {
    if (lumMod == 1f && lumOff == 0f) return color
    fun ch(c: Int): Int = ((c * lumMod + 255f * lumOff).toInt()).coerceIn(0, 255)
    val a = (color ushr 24) and 0xFF
    val r = ch((color shr 16) and 0xFF)
    val g = ch((color shr 8) and 0xFF)
    val b = ch(color and 0xFF)
    return (a shl 24) or (r shl 16) or (g shl 8) or b
}
