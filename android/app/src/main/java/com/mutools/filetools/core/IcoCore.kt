package com.mutools.filetools.core

import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.RectF

/**
 * 「图片转 ICO」的位图处理核心，逐条对应桌面版 core/iconizer.py：
 * 降采样 → 正方形化 → 逐档缩放加圆角（4 倍超采样抗锯齿）→ 打包七档尺寸。
 */
object IcoCore {

    val ICO_SIZES = intArrayOf(256, 128, 64, 48, 32, 24, 16)
    const val SOURCE_MAX = 1024
    private const val SS = 4  // 圆角蒙版超采样倍数

    /** 等比缩小到最大边长以内（只缩小，不放大）。 */
    fun downscale(src: Bitmap, maxSide: Int = SOURCE_MAX): Bitmap {
        val longest = maxOf(src.width, src.height)
        if (longest <= maxSide) return src
        val scale = maxSide.toFloat() / longest
        val w = maxOf(1, (src.width * scale).toInt())
        val h = maxOf(1, (src.height * scale).toInt())
        return Bitmap.createScaledBitmap(src, w, h, true)
    }

    /** 变成正方形：crop=true 居中裁剪，否则用透明像素补边。 */
    fun toSquare(src: Bitmap, crop: Boolean): Bitmap {
        if (src.width == src.height) return src
        return if (crop) {
            val side = minOf(src.width, src.height)
            val left = (src.width - side) / 2
            val top = (src.height - side) / 2
            Bitmap.createBitmap(src, left, top, side, side)
        } else {
            val side = maxOf(src.width, src.height)
            val out = Bitmap.createBitmap(side, side, Bitmap.Config.ARGB_8888)
            Canvas(out).drawBitmap(
                src,
                ((side - src.width) / 2).toFloat(),
                ((side - src.height) / 2).toFloat(),
                null
            )
            out
        }
    }

    /** 正方形图缩放到指定边长（允许放大，小尺寸档位需要）。 */
    fun fitSquare(src: Bitmap, side: Int): Bitmap =
        if (src.width == side) src else Bitmap.createScaledBitmap(src, side, side, true)

    /**
     * 给图加圆角：圆角外透明；图片自身已有的透明区域保持不变
     * （alpha 逐像素取暗合并，对应 ImageChops.darker）。
     */
    fun applyRound(src: Bitmap, radiusPct: Float): Bitmap {
        if (radiusPct <= 0f) return src
        val square = if (src.width != src.height) toSquare(src, true) else src
        val side = square.width
        val mask = roundedMask(side, radiusPct)
        val out = square.copy(Bitmap.Config.ARGB_8888, false)
        val n = side * side
        val px = IntArray(n)
        val mp = IntArray(n)
        out.getPixels(px, 0, side, 0, 0, side, side)
        mask.getPixels(mp, 0, side, 0, 0, side, side)
        for (i in 0 until n) {
            val a = (px[i] ushr 24) and 0xFF
            val m = (mp[i] ushr 24) and 0xFF
            px[i] = (minOf(a, m) shl 24) or (px[i] and 0x00FFFFFF)
        }
        out.setPixels(px, 0, side, 0, 0, side, side)
        mask.recycle()
        return out
    }

    /** 抗锯齿圆角蒙版：按 4 倍超采样绘制圆角矩形后缩小回目标尺寸。 */
    private fun roundedMask(size: Int, radiusPct: Float): Bitmap {
        val pct = radiusPct.coerceIn(0f, 50f)
        val big = size * SS
        val mask = Bitmap.createBitmap(big, big, Bitmap.Config.ARGB_8888)
        val c = Canvas(mask)
        val p = Paint(Paint.ANTI_ALIAS_FLAG).apply {
            color = Color.WHITE
            style = Paint.Style.FILL
        }
        val r = big * pct / 100f
        c.drawRoundRect(RectF(0f, 0f, big.toFloat(), big.toFloat()), r, r, p)
        return Bitmap.createScaledBitmap(mask, size, size, true)
    }

    /** 每一档尺寸分别缩放、加圆角，保证小尺寸的圆角比例一致、边缘清晰。 */
    fun buildFrames(square: Bitmap, radiusPct: Float): List<IcoEncoder.Frame> {
        val side = square.width
        require(side == square.height) { "输入必须是正方形" }
        return ICO_SIZES.map { size ->
            val b = applyRound(fitSquare(square, size), radiusPct)
            val px = IntArray(size * size)
            b.getPixels(px, 0, size, 0, 0, size, size)
            if (b !== square) b.recycle()
            IcoEncoder.Frame(size, px)
        }
    }

    /** 灰白棋盘格底图，衬托透明区域（对应 iconizer.on_checkerboard）。 */
    fun onCheckerboard(icon: Bitmap, cell: Int = 8): Bitmap {
        val out = Bitmap.createBitmap(icon.width, icon.height, Bitmap.Config.ARGB_8888)
        val c = Canvas(out)
        c.drawColor(Color.WHITE)
        val p = Paint().apply { color = 0xFFCCCCCC.toInt() }
        var y = 0
        var row = 0
        while (y < icon.height) {
            var x = 0
            var col = 0
            while (x < icon.width) {
                if ((col + row) % 2 == 1) {
                    c.drawRect(
                        x.toFloat(), y.toFloat(),
                        minOf(x + cell, icon.width).toFloat(),
                        minOf(y + cell, icon.height).toFloat(),
                        p
                    )
                }
                x += cell
                col++
            }
            y += cell
            row++
        }
        c.drawBitmap(icon, 0f, 0f, null)
        return out
    }
}
