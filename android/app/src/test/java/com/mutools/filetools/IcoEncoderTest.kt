package com.mutools.filetools

import com.mutools.filetools.core.IcoEncoder
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test

/**
 * ICO 编码器无界面自测：结构字节逐一校验（对齐桌面版 tests/test_iconizer.py 的思路）。
 */
class IcoEncoderTest {

    private fun frame(size: Int, argb: Int): IcoEncoder.Frame =
        IcoEncoder.Frame(size, IntArray(size * size) { argb })

    @Test
    fun `目录与条目结构正确`() {
        val bytes = IcoEncoder.encode(listOf(frame(16, 0xFF336699.toInt()), frame(256, 0x80FF0000.toInt())))
        assertEquals(0, bytes[0].toInt()); assertEquals(0, bytes[1].toInt())   // reserved
        assertEquals(1, bytes[2].toInt()); assertEquals(0, bytes[3].toInt())   // type=icon
        assertEquals(2, bytes[4].toInt()); assertEquals(0, bytes[5].toInt())   // count=2

        // 条目 0（256px，按尺寸降序在前）：宽高记 0
        assertEquals(0, bytes[6].toInt())
        assertEquals(0, bytes[7].toInt())

        // 条目 1（16px）：宽 16 高 16，planes 1、bpp 32
        assertEquals(16, bytes[6 + 16].toInt())
        assertEquals(16, bytes[7 + 16].toInt())
        assertEquals(1, bytes[6 + 16 + 4].toInt())   // planes 低字节
        assertEquals(32, bytes[6 + 16 + 6].toInt())  // bpp 低字节
    }

    @Test
    fun `偏移与总长一致`() {
        val frames = listOf(frame(16, 0xFF336699.toInt()), frame(256, 0xFF000000.toInt()))
        val bytes = IcoEncoder.encode(frames)
        fun u32(o: Int): Long =
            (bytes[o].toLong() and 0xFF) or ((bytes[o + 1].toLong() and 0xFF) shl 8) or
                ((bytes[o + 2].toLong() and 0xFF) shl 16) or ((bytes[o + 3].toLong() and 0xFF) shl 24)

        val offset0 = u32(6 + 8 + 4 + 0)     // entry0 dwImageOffset 在 entry 起始 +12
        val size0 = u32(6 + 8)
        val offset1 = u32(6 + 16 + 12)
        val size1 = u32(6 + 16 + 8)
        assertEquals(6 + 16 * 2L, offset0)
        assertEquals(offset0 + size0, offset1)
        assertEquals(offset1 + size1, bytes.size.toLong())
    }

    @Test
    fun `BMP 档底部像素可回读且为 BGRA 布局`() {
        // 顶部一行为红 0xFFFF0000，底部一行为蓝 0xFF0000FF
        val size = 8
        val px = IntArray(size * size) { 0xFF00FF00.toInt() } // 绿色填充
        for (x in 0 until size) {
            px[x] = 0xFFFF0000.toInt()
            px[(size - 1) * size + x] = 0xFF0000FF.toInt()
        }
        val bytes = IcoEncoder.encode(listOf(IcoEncoder.Frame(size, px)))
        val dataOffset = 6 + 16
        assertEquals(40, bytes[dataOffset].toInt())                       // biSize=40
        // biHeight = 16（8 × 2），小端在 offset dataOffset+8
        assertEquals(16, bytes[dataOffset + 8].toInt())
        assertEquals(32, bytes[dataOffset + 14].toInt())                  // biBitCount=32

        // XOR 数据自底向上：第一行像素（文件 offset dataOffset+40）是图像最后一行（蓝）
        val firstPixel = dataOffset + 40
        val b = bytes[firstPixel].toInt() and 0xFF
        val g = bytes[firstPixel + 1].toInt() and 0xFF
        val r = bytes[firstPixel + 2].toInt() and 0xFF
        val a = bytes[firstPixel + 3].toInt() and 0xFF
        assertEquals(0xFF, a); assertEquals(0x00, r); assertEquals(0x00, g); assertEquals(0xFF, b)
    }

    @Test
    fun `256 档使用 PNG 存储`() {
        val bytes = IcoEncoder.encode(listOf(frame(256, 0xFF336699.toInt())))
        val sig = intArrayOf(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
        val start = 6 + 16
        for (i in sig.indices) assertEquals(sig[i], bytes[start + i].toInt() and 0xFF)
        // IHDR 宽度 256（大端）在 PNG 偏移 16
        val wOff = start + 16
        assertEquals(0, bytes[wOff].toInt())
        assertEquals(1, bytes[wOff + 2].toInt())  // 256 = 0x00 00 01 00
    }

    @Test
    fun `单帧也能编码`() {
        val bytes = IcoEncoder.encode(listOf(frame(32, 0xFF336699.toInt())))
        assertEquals(1, bytes[4].toInt())
        assertTrue(bytes.size > 6 + 16)
    }
}
