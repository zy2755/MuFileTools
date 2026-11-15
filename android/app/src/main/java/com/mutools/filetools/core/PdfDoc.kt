package com.mutools.filetools.core

import android.content.Context
import android.graphics.Bitmap
import android.graphics.Color
import android.graphics.pdf.PdfRenderer
import android.net.Uri
import com.tom_roush.pdfbox.pdmodel.PDDocument
import java.io.Closeable
import java.io.OutputStream

/**
 * 一个打开的 PDF：缩略图 / 按 DPI 渲染 / 页面提取（pdfbox 复制保留矢量）。
 * 对应桌面版 core/sources.py 的 PdfSource。
 */
class PdfDoc(private val context: Context, val uri: Uri) : PageSource {

    private val pfd = context.contentResolver.openFileDescriptor(uri, "r")!!
    private val renderer = PdfRenderer(pfd)
    private val lock = Any()

    override val pageCount: Int
        get() = synchronized(lock) { renderer.pageCount }

    override val unitName: String get() = "页"

    override fun label(index: Int): String = "第 ${index + 1} 页"

    override fun fileLabel(index: Int): String = (index + 1).toString()

    /** 渲染一页，宽为 targetWidth（等比），白底，用于缩略图与预览。 */
    override fun renderPage(index: Int, targetWidth: Int): Bitmap = synchronized(lock) {
        val page = renderer.openPage(index)
        try {
            val w = targetWidth.coerceAtLeast(1)
            val h = maxOf(1, (page.height.toLong() * w / page.width).toInt())
            val bmp = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888)
            bmp.eraseColor(Color.WHITE)
            page.render(bmp, null, null, PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY)
            bmp
        } finally {
            page.close()
        }
    }

    /** 按 DPI 渲染一页（PDF 页面单位是 1/72 英寸），白底，用于导出。 */
    override fun renderPageDpi(index: Int, dpi: Int): Bitmap = synchronized(lock) {
        val page = renderer.openPage(index)
        try {
            val scale = dpi / 72f
            var w = maxOf(1, (page.width * scale).toInt())
            var h = maxOf(1, (page.height * scale).toInt())
            // 单页像素上限约 4000 万：高 DPI 大页面整体缩档，防 OOM
            if (w.toLong() * h > MAX_PIXELS) {
                val f = kotlin.math.sqrt(MAX_PIXELS.toDouble() / (w.toLong() * h))
                w = maxOf(1, (w * f).toInt())
                h = maxOf(1, (h * f).toInt())
            }
            val bmp = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888)
            bmp.eraseColor(Color.WHITE)
            page.render(bmp, null, null, PdfRenderer.Page.RENDER_MODE_FOR_PRINT)
            bmp
        } finally {
            page.close()
        }
    }

    /**
     * 把选中页写成新 PDF：重新加载原文档、倒序删除未选页后保存，
     * 保留矢量、文字、链接与书签（对应桌面版 PdfSource.export_pdf 的保真做法）。
     * dpi 参数在这里用不上（不重新光栅化），保留是为了接口一致。
     */
    override fun exportPdf(selected: IntArray, dpi: Int, out: OutputStream) {
        require(selected.isNotEmpty()) { "没有选中任何页面" }
        val src = context.contentResolver.openInputStream(uri)!!
        PDDocument.load(src).use { doc ->
            val keep = selected.toHashSet()
            for (i in doc.numberOfPages - 1 downTo 0) {
                if (i !in keep) doc.removePage(doc.getPage(i))
            }
            doc.save(out)
        }
    }

    override fun close() {
        synchronized(lock) {
            renderer.close()
            pfd.close()
        }
    }

    companion object {
        private const val MAX_PIXELS = 40_000_000L
    }
}
