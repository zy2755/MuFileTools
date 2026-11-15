package com.mutools.filetools.core

import android.content.Context
import android.graphics.Bitmap
import android.net.Uri
import androidx.documentfile.provider.DocumentFile

/** SAF 输出辅助：文件名清洗、去重命名、写文件。 */
object Outputs {

    private val BAD = Regex("[\\\\/:*?\"<>|\\r\\n\\t]+")

    /** 文件名里的非法字符替换为下划线（对应桌面版 exporter._clean）。 */
    fun clean(s: String): String = s.replace(BAD, "_").trim()

    fun stem(name: String): String = name.substringBeforeLast('.')

    fun mimeOf(name: String): String = when (name.substringAfterLast('.').lowercase()) {
        "png" -> "image/png"
        "jpg", "jpeg" -> "image/jpeg"
        "pdf" -> "application/pdf"
        "ico" -> "image/x-icon"
        else -> "application/octet-stream"
    }

    /** 目录内重名时自动追加 (1)、(2)…（桌面版默认 rename 策略）。 */
    fun uniqueIn(dir: DocumentFile, name: String): DocumentFile? {
        if (dir.findFile(name) == null) return dir.createFile(mimeOf(name), name)
        val stem = stem(name)
        val ext = name.substringAfterLast('.')
        for (i in 1..999) {
            val n = "$stem($i).$ext"
            if (dir.findFile(n) == null) return dir.createFile(mimeOf(name), n)
        }
        return null
    }

    fun writeBytes(ctx: Context, uri: Uri, bytes: ByteArray) {
        ctx.contentResolver.openOutputStream(uri, "wt")?.use { it.write(bytes) }
    }

    /** png=false 时按 JPEG quality 压缩；位图须为无透明（白底）内容。 */
    fun writeBitmap(ctx: Context, uri: Uri, bmp: Bitmap, png: Boolean, quality: Int = 92): Boolean {
        val out = ctx.contentResolver.openOutputStream(uri, "wt") ?: return false
        out.use {
            return bmp.compress(
                if (png) Bitmap.CompressFormat.PNG else Bitmap.CompressFormat.JPEG,
                quality,
                it
            )
        }
    }
}
