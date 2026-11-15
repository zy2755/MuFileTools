package com.mutools.filetools.core

/**
 * 复合文档（OLE2 / CFB）读取器 —— .xls / .ppt / .doc 的外壳都是它。
 *
 * 结构（MS-CFB）：
 *   512 字节头 → FAT 链（记录「下一个扇区号」的数组）
 *            → 目录流（一串 128 字节的条目，像个小文件系统）
 *            → 每个流按 FAT 链把扇区串起来
 * 小于 4096 字节的流不占普通扇区，而是塞在根条目那条「迷你流」里，
 * 用另一个更小的 64 字节扇区表（miniFAT）寻址 —— 坑就在这，漏了它小文件读不出来。
 *
 * 只做读取，且只读到 ByteArray：.xls 的 Workbook 流通常几十 KB 到几 MB，
 * 整份读进内存比维护随机读的流对象简单得多。
 */
internal class Cfb private constructor(private val data: ByteArray) {

    private var sectorSize = 512
    private var miniSectorSize = 64
    private var miniCutoff = 4096

    private var fat = IntArray(0)
    private var miniFat = IntArray(0)
    private var miniStream = ByteArray(0)

    private class Entry(
        val name: String,
        val type: Int,
        val start: Int,
        val size: Long,
    )

    private val root = ArrayList<Entry>()      // 目录条目（扁平化，够用）
    private val byName = HashMap<String, Entry>()

    init {
        parseHeader()
        parseFat()
        parseDirectory()
    }

    /** 按名字取流（大小写不敏感）。找不到返回 null。 */
    fun stream(name: String): ByteArray? {
        val e = byName[name.lowercase()] ?: return null
        if (e.type != TYPE_STREAM) return null
        val size = e.size.toInt()
        if (size <= 0) return ByteArray(0)
        return if (e.size < miniCutoff) readMini(e.start, size) else readChain(e.start, size)
    }

    fun hasStream(name: String): Boolean = byName[name.lowercase()]?.type == TYPE_STREAM

    // ---------------------------------------------------------------- 头部

    private fun parseHeader() {
        require(data.size >= 512) { "不是复合文档（文件太小）" }
        // 魔数 D0 CF 11 E0 A1 B1 1A E1，小端读成两个 32 位值
        require(u32(0) == 0xE011CFD0.toInt() && u32(4) == 0xE11AB1A1.toInt()) {
            "不是复合文档（魔数不符）"
        }

        sectorSize = 1 shl u16(0x1E)
        miniSectorSize = 1 shl u16(0x20)
        miniCutoff = u32(0x38)

        val numFatSectors = u32(0x2C)
        firstDirSector = i32(0x30)
        firstMiniFatSector = i32(0x3C)
        numMiniFatSectors = u32(0x40)
        firstDifatSector = i32(0x44)
        numDifatSectors = u32(0x48)

        // 头部里能放 109 个 FAT 扇区号，超出的部分挂在 DIFAT 链上
        val fatSectors = ArrayList<Int>()
        for (i in 0 until 109) {
            val s = i32(0x4C + i * 4)
            if (s >= 0) fatSectors += s
        }
        var difat = firstDifatSector
        var guard = 0
        while (difat >= 0 && guard++ < 1 shl 16 && fatSectors.size < numFatSectors) {
            val base = offsetOf(difat)
            if (base < 0 || base + sectorSize > data.size) break
            val per = sectorSize / 4 - 1
            for (i in 0 until per) {
                val s = i32(base + i * 4)
                if (s >= 0) fatSectors += s
            }
            difat = i32(base + per * 4)
        }
        this.fatSectorList = fatSectors.toIntArray()
    }

    private var firstDirSector = -1
    private var firstMiniFatSector = -1
    private var numMiniFatSectors = 0
    private var firstDifatSector = -1
    private var numDifatSectors = 0
    private var fatSectorList = IntArray(0)

    /** 扇区号 → 文件内偏移。注意 +1：头占了一整个扇区。 */
    private fun offsetOf(sector: Int): Int = (sector + 1) * sectorSize

    // ---------------------------------------------------------------- FAT

    private fun parseFat() {
        val per = sectorSize / 4
        val table = IntArray(fatSectorList.size * per)
        var n = 0
        for (s in fatSectorList) {
            val base = offsetOf(s)
            if (base < 0 || base + sectorSize > data.size) continue
            for (i in 0 until per) table[n++] = i32(base + i * 4)
        }
        fat = if (n == table.size) table else table.copyOf(n)

        if (firstMiniFatSector >= 0 && numMiniFatSectors > 0) {
            val bytes = readChain(firstMiniFatSector, numMiniFatSectors * sectorSize)
            val cnt = bytes.size / 4
            val mf = IntArray(cnt)
            for (i in 0 until cnt) mf[i] = le32(bytes, i * 4)
            miniFat = mf
        }
    }

    /** 顺着 FAT 链把扇区拼成连续字节。 */
    private fun readChain(start: Int, size: Int): ByteArray {
        val out = ByteArray(size)
        var pos = 0
        var s = start
        var guard = 0
        while (s >= 0 && pos < size && guard++ < 1 shl 20) {
            val base = offsetOf(s)
            if (base < 0 || base >= data.size) break
            val n = minOf(sectorSize, size - pos, data.size - base)
            if (n <= 0) break
            System.arraycopy(data, base, out, pos, n)
            pos += n
            s = if (s < fat.size) fat[s] else -1
        }
        return if (pos == size) out else out.copyOf(pos)
    }

    /** 迷你流里的链（64 字节一扇区）。 */
    private fun readMini(start: Int, size: Int): ByteArray {
        val out = ByteArray(size)
        var pos = 0
        var s = start
        var guard = 0
        while (s >= 0 && pos < size && guard++ < 1 shl 20) {
            val base = s * miniSectorSize
            if (base < 0 || base >= miniStream.size) break
            val n = minOf(miniSectorSize, size - pos, miniStream.size - base)
            if (n <= 0) break
            System.arraycopy(miniStream, base, out, pos, n)
            pos += n
            s = if (s < miniFat.size) miniFat[s] else -1
        }
        return if (pos == size) out else out.copyOf(pos)
    }

    // ---------------------------------------------------------------- 目录

    private fun parseDirectory() {
        if (firstDirSector < 0) return
        // 目录流长度未知：先按链一直读到底
        val dirBytes = readChainToEnd(firstDirSector)
        var i = 0
        while (i + 128 <= dirBytes.size) {
            val nameLen = le16(dirBytes, i + 0x40)
            val type = dirBytes[i + 0x42].toInt() and 0xFF
            val start = le32(dirBytes, i + 0x74)
            val sizeLow = le32(dirBytes, i + 0x78).toLong() and 0xFFFFFFFFL
            val sizeHigh = le32(dirBytes, i + 0x7C).toLong() and 0xFFFFFFFFL
            val size = sizeLow or (sizeHigh shl 32)
            if (type in 1..5 && nameLen in 2..64) {
                val chars = (nameLen - 2) / 2
                val sb = StringBuilder(chars)
                for (k in 0 until chars) sb.append(le16(dirBytes, i + k * 2).toChar())
                val e = Entry(sb.toString(), type, start, size)
                root += e
                byName[e.name.lowercase()] = e
                if (type == TYPE_ROOT) miniStream = readChain(start, minOf(size, Int.MAX_VALUE.toLong()).toInt())
            }
            i += 128
        }
    }

    private fun readChainToEnd(start: Int): ByteArray {
        val out = java.io.ByteArrayOutputStream()
        var s = start
        var guard = 0
        while (s >= 0 && guard++ < 1 shl 20) {
            val base = offsetOf(s)
            if (base < 0 || base >= data.size) break
            val n = minOf(sectorSize, data.size - base)
            if (n <= 0) break
            out.write(data, base, n)
            s = if (s < fat.size) fat[s] else -1
        }
        return out.toByteArray()
    }

    // ---------------------------------------------------------------- 小工具

    private fun u16(off: Int): Int = le16(data, off)
    private fun u32(off: Int): Int = le32(data, off)
    private fun i32(off: Int): Int = le32(data, off)

    companion object {
        private const val TYPE_STREAM = 2
        private const val TYPE_ROOT = 5

        fun le16(b: ByteArray, off: Int): Int =
            (b[off].toInt() and 0xFF) or ((b[off + 1].toInt() and 0xFF) shl 8)

        fun le32(b: ByteArray, off: Int): Int =
            (b[off].toInt() and 0xFF) or
                ((b[off + 1].toInt() and 0xFF) shl 8) or
                ((b[off + 2].toInt() and 0xFF) shl 16) or
                ((b[off + 3].toInt() and 0xFF) shl 24)

        fun open(bytes: ByteArray): Cfb = Cfb(bytes)
    }
}
