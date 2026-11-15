package com.mutools.filetools.core

import java.io.ByteArrayOutputStream
import java.util.zip.CRC32
import java.util.zip.Deflater

/**
 * ICO 打包器（纯 JVM，无 Android 依赖，可直接单元测试）。
 *
 * 按 Windows 规范写 .ico：
 * - 256 档用 PNG 压缩存储（Vista+ 推荐，体积小）；
 * - 其余档用 32 位 BMP（BITMAPINFOHEADER 高度为两倍 + AND 掩码），兼容所有资源管理器。
 *
 * 文件结构：ICONDIR(6B) + N 个 ICONDIRENTRY(16B) + 各档图像数据。
 */
object IcoEncoder {

    /** 一档图像：ARGB 像素（行优先、非预乘），size 必须与 sqrt(pixels.size) 一致。 */
    class Frame(val size: Int, val pixels: IntArray) {
        init {
            require(pixels.size == size * size) { "像素数 ${pixels.size} != ${size}x${size}" }
        }
    }

    fun encode(frames: List<Frame>): ByteArray {
        require(frames.isNotEmpty()) { "没有图像帧" }
        val ordered = frames.sortedByDescending { it.size }
        val blobs = ordered.map { encodeImage(it) }

        val out = ByteArrayOutputStream()
        writeU16(out, 0)              // reserved
        writeU16(out, 1)              // type: 1 = icon
        writeU16(out, ordered.size)
        var offset = 6 + 16 * ordered.size
        for (i in ordered.indices) {
            val f = ordered[i]
            out.write(if (f.size >= 256) 0 else f.size)
            out.write(if (f.size >= 256) 0 else f.size)
            out.write(0)              // 调色板色数
            out.write(0)              // reserved
            writeU16(out, 1)          // planes
            writeU16(out, 32)         // bpp
            writeU32(out, blobs[i].size)
            writeU32(out, offset)
            offset += blobs[i].size
        }
        for (b in blobs) out.write(b)
        return out.toByteArray()
    }

    private fun encodeImage(f: Frame): ByteArray =
        if (f.size >= 256) pngBytes(f) else bmpBytes(f)

    // ---------------------------------------------------------------- BMP 档

    private fun bmpBytes(f: Frame): ByteArray {
        val maskRow = (f.size + 31) / 32 * 4
        val xorSize = f.size * f.size * 4
        val maskSize = maskRow * f.size
        val out = ByteArrayOutputStream(xorSize + maskSize + 40)

        // BITMAPINFOHEADER（40 字节，小端）
        writeU32(out, 40)
        writeU32(out, f.size)         // biWidth
        writeU32(out, f.size * 2)     // biHeight = 图像高度 × 2（XOR + AND）
        writeU16(out, 1)              // biPlanes
        writeU16(out, 32)             // biBitCount
        writeU32(out, 0)              // biCompression = BI_RGB
        writeU32(out, xorSize + maskSize)
        writeU32(out, 0)
        writeU32(out, 0)
        writeU32(out, 0)
        writeU32(out, 0)

        // XOR 数据：BGRA，自底向上
        for (y in f.size - 1 downTo 0) {
            for (x in 0 until f.size) {
                val argb = f.pixels[y * f.size + x]
                out.write(argb and 0xFF)          // B
                out.write((argb shr 8) and 0xFF)  // G
                out.write((argb shr 16) and 0xFF) // R
                out.write((argb shr 24) and 0xFF) // A
            }
        }
        // AND 掩码：有 alpha 通道时全 0
        out.write(ByteArray(maskSize))
        return out.toByteArray()
    }

    // ---------------------------------------------------------------- PNG 档

    private fun pngBytes(f: Frame): ByteArray {
        val out = ByteArrayOutputStream()
        out.write(byteArrayOf(0x89.toByte(), 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A), 0, 8)

        val ihdr = ByteArrayOutputStream()
        writeU32be(ihdr, f.size)
        writeU32be(ihdr, f.size)
        ihdr.write(8)   // bit depth
        ihdr.write(6)   // color type: RGBA
        ihdr.write(0)   // compression
        ihdr.write(0)   // filter
        ihdr.write(0)   // interlace
        chunk(out, "IHDR", ihdr.toByteArray())

        // 逐行扫描，每行前置 filter byte 0
        val stride = 1 + f.size * 4
        val raw = ByteArray(f.size * stride)
        var p = 0
        for (y in 0 until f.size) {
            raw[p++] = 0
            for (x in 0 until f.size) {
                val argb = f.pixels[y * f.size + x]
                raw[p++] = (argb shr 16).toByte()  // R
                raw[p++] = (argb shr 8).toByte()   // G
                raw[p++] = argb.toByte()           // B
                raw[p++] = (argb shr 24).toByte()  // A
            }
        }
        val def = Deflater()
        def.setInput(raw)
        def.finish()
        val comp = ByteArrayOutputStream()
        val buf = ByteArray(8192)
        while (!def.finished()) {
            val n = def.deflate(buf)
            if (n > 0) comp.write(buf, 0, n)
        }
        def.end()
        chunk(out, "IDAT", comp.toByteArray())
        chunk(out, "IEND", ByteArray(0))
        return out.toByteArray()
    }

    private fun chunk(out: ByteArrayOutputStream, type: String, data: ByteArray) {
        writeU32be(out, data.size)
        val typeBytes = type.toByteArray(Charsets.US_ASCII)
        out.write(typeBytes, 0, typeBytes.size)
        out.write(data, 0, data.size)
        val crc = CRC32()
        crc.update(typeBytes)
        crc.update(data)
        writeU32be(out, crc.value.toInt())
    }

    // ---------------------------------------------------------------- 基础写出

    private fun writeU16(out: ByteArrayOutputStream, v: Int) {
        out.write(v and 0xFF)
        out.write((v shr 8) and 0xFF)
    }

    private fun writeU32(out: ByteArrayOutputStream, v: Int) {
        out.write(v and 0xFF)
        out.write((v shr 8) and 0xFF)
        out.write((v shr 16) and 0xFF)
        out.write((v shr 24) and 0xFF)
    }

    private fun writeU32be(out: ByteArrayOutputStream, v: Int) {
        out.write((v ushr 24) and 0xFF)
        out.write((v ushr 16) and 0xFF)
        out.write((v ushr 8) and 0xFF)
        out.write(v and 0xFF)
    }
}
