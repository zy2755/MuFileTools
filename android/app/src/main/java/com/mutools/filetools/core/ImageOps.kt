package com.mutools.filetools.core

import android.content.Context
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.ImageDecoder
import android.graphics.Matrix
import android.net.Uri
import android.os.Build
import androidx.exifinterface.media.ExifInterface

/** 图片读取（EXIF 纠偏）与长图拼接。 */
object ImageOps {

    /**
     * 读取图片并纠正 EXIF 方向，最长边压到 maxSide 以内（只缩小，不放大）。
     * BitmapFactory 失败时回退 ImageDecoder（API 28+，AVIF/HEIC/特殊 WebP 更稳）。
     */
    fun decodeUpright(ctx: Context, uri: Uri, maxSide: Int): Bitmap? {
        val bmp = decodeByBitmapFactory(ctx, uri, maxSide)
            ?: decodeByImageDecoder(ctx, uri, maxSide)
            ?: return null
        val exifStream = try {
            ctx.contentResolver.openInputStream(uri)
        } catch (_: Exception) {
            null
        }
        val rotated = rotateByExif(exifStream, bmp)
        return IcoCore.downscale(rotated, maxSide)
    }

    private fun decodeByBitmapFactory(ctx: Context, uri: Uri, maxSide: Int): Bitmap? {
        val resolver = ctx.contentResolver
        val bounds = BitmapFactory.Options().apply { inJustDecodeBounds = true }
        val boundsOk = try {
            resolver.openInputStream(uri)?.use { BitmapFactory.decodeStream(it, null, bounds) } != null
        } catch (_: Exception) {
            false
        }
        if (!boundsOk || bounds.outWidth <= 0 || bounds.outHeight <= 0) return null

        var sample = 1
        val longest = maxOf(bounds.outWidth, bounds.outHeight)
        while (longest / (sample * 2) >= maxSide) sample *= 2

        val opts = BitmapFactory.Options().apply {
            inSampleSize = sample
            inPreferredConfig = Bitmap.Config.ARGB_8888
        }
        return try {
            resolver.openInputStream(uri)?.use { BitmapFactory.decodeStream(it, null, opts) }
        } catch (_: Exception) {
            null
        }
    }

    /** ImageDecoder 后备：注意必须 setMutableRequired(true)，否则拿到硬件位图无法 getPixels。 */
    private fun decodeByImageDecoder(ctx: Context, uri: Uri, maxSide: Int): Bitmap? {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.P) return null
        return try {
            val source = ImageDecoder.createSource(ctx.contentResolver, uri)
            ImageDecoder.decodeBitmap(source) { decoder, info, _ ->
                decoder.isMutableRequired = true
                val longest = maxOf(info.size.width, info.size.height)
                var sample = 1
                while (longest / (sample * 2) >= maxSide) sample *= 2
                if (sample > 1) decoder.setTargetSampleSize(sample)
            }
        } catch (_: Exception) {
            null
        }
    }

    private fun rotateByExif(stream: java.io.InputStream?, bmp: Bitmap): Bitmap {
        if (stream == null) return bmp
        val orientation = try {
            stream.use { ExifInterface(it).getAttributeInt(ExifInterface.TAG_ORIENTATION, ExifInterface.ORIENTATION_NORMAL) }
        } catch (_: Exception) {
            ExifInterface.ORIENTATION_NORMAL
        }
        val m = Matrix()
        when (orientation) {
            ExifInterface.ORIENTATION_ROTATE_90 -> m.postRotate(90f)
            ExifInterface.ORIENTATION_ROTATE_180 -> m.postRotate(180f)
            ExifInterface.ORIENTATION_ROTATE_270 -> m.postRotate(270f)
            ExifInterface.ORIENTATION_FLIP_HORIZONTAL -> m.preScale(-1f, 1f)
            ExifInterface.ORIENTATION_FLIP_VERTICAL -> m.preScale(1f, -1f)
            ExifInterface.ORIENTATION_TRANSPOSE -> { m.postRotate(90f); m.postScale(-1f, 1f) }
            ExifInterface.ORIENTATION_TRANSVERSE -> { m.postRotate(270f); m.postScale(-1f, 1f) }
        }
        return if (m.isIdentity) bmp else Bitmap.createBitmap(bmp, 0, 0, bmp.width, bmp.height, m, true)
    }

    /**
     * 竖版长图：统一宽度（取最宽页）后顺序拼接，白底；
     * 总高超上限时整体缩小（JPEG 单边 ≤65535 的约束）。
     */
    fun stitch(pages: List<Bitmap>, maxTotalHeight: Int = 65000): Bitmap {
        require(pages.isNotEmpty()) { "没有可拼接的页面" }
        val maxW = pages.maxOf { it.width }
        fun heightFor(w: Int, p: Bitmap) = maxOf(1, ((p.height.toLong() * w) / p.width).toInt())
        var targetW = maxW
        var totalH = pages.sumOf { heightFor(targetW, it).toLong() }.coerceAtMost(Int.MAX_VALUE.toLong()).toInt()
        if (totalH > maxTotalHeight) {
            targetW = maxOf(1, ((targetW.toLong() * maxTotalHeight) / totalH).toInt())
            totalH = pages.sumOf { heightFor(targetW, it).toLong() }.coerceAtMost(Int.MAX_VALUE.toLong()).toInt()
        }
        val out = Bitmap.createBitmap(targetW, totalH, Bitmap.Config.ARGB_8888)
        val c = Canvas(out)
        c.drawColor(Color.WHITE)
        var y = 0
        for (p in pages) {
            val h = heightFor(targetW, p)
            val scaled = if (p.width == targetW) p else Bitmap.createScaledBitmap(p, targetW, h, true)
            c.drawBitmap(scaled, 0f, y.toFloat(), null)
            y += h
            if (scaled !== p) scaled.recycle()
        }
        return out
    }
}
