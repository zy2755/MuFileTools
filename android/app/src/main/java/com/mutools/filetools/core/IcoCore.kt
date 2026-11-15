package com.mutools.filetools.core

import android.graphics.Bitmap
import android.graphics.BitmapShader
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.RectF
import android.graphics.Shader

/**
 * 「图片转 ICO」的位图处理核心，逐条对应桌面版 core/iconizer.py：
 * 降采样 → 正方形化 → 逐档缩放加圆角（4 倍超采样抗锯齿）→ 打包七档尺寸。
 */
object IcoCore {

    val ICO_SIZES = intArrayOf(256, 128, 64, 48, 32, 24, 16)
    const val SOURCE_MAX = 1024

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
     * 给图加圆角：圆角外透明；图片自身已有的透明区域保持不变。
     *
     * v1.1.3 之前用 `copy() + setPixels 改 alpha`：`copy()` 会把源位图的 hasAlpha
     * 一起复制过去，而相册/相机出来的 JPEG 位图 hasAlpha=false —— 这时写进像素的
     * alpha 在绘制阶段被整张忽略，圆角"加了但看不见"（白底图尤其明显：四角照旧是白的）。
     * 现在改为新建 ARGB 位图 + BitmapShader 填进圆角矩形：圆角外根本不绘制，
     * 新建位图本就是全透明，既不依赖源图的 hasAlpha，也不用逐像素数组（省 8MB 临时内存）。
     */
    fun applyRound(src: Bitmap, radiusPct: Float): Bitmap {
        if (radiusPct <= 0f) return src
        val square = if (src.width != src.height) toSquare(src, true) else src
        val side = square.width
        val pct = radiusPct.coerceIn(0f, 50f)
        val r = side * pct / 100f
        val out = Bitmap.createBitmap(side, side, Bitmap.Config.ARGB_8888)
        out.setHasAlpha(true)
        val c = Canvas(out)
        val paint = Paint(Paint.ANTI_ALIAS_FLAG or Paint.FILTER_BITMAP_FLAG)
        // 用 BitmapShader 把原图「填」进圆角矩形：圆角外根本没有绘制，
        // 新建位图那里本来就是全透明，不依赖 xfermode 去削 alpha
        paint.shader = BitmapShader(square, Shader.TileMode.CLAMP, Shader.TileMode.CLAMP)
        c.drawRoundRect(RectF(0f, 0f, side.toFloat(), side.toFloat()), r, r, paint)
        return out
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
