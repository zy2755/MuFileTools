package com.mutools.filetools.core

import android.content.Context
import android.graphics.Bitmap
import android.graphics.Color
import android.graphics.pdf.PdfDocument
import android.net.Uri
import java.io.Closeable
import java.io.OutputStream

/**
 * 「页面提取」的单位来源。三种格式的「一页」含义不同：
 *   PDF → 页（page）  PPT → 幻灯片（slide）  Excel → 工作表（sheet）
 * 统一抽象成 pageCount + renderPage，界面层就不用关心底下是什么格式。
 */
interface PageSource : Closeable {

    val pageCount: Int

    /** 单位名，用于「已选 X / Y 页」这类文案。 */
    val unitName: String

    /** 网格角标：第 3 页 / 销售明细。 */
    fun label(index: Int): String

    /** 导出文件名用的短标签：3 / 销售明细。 */
    fun fileLabel(index: Int): String

    /** 缩略图：宽固定为 targetWidth，等比高。 */
    fun renderPage(index: Int, targetWidth: Int): Bitmap

    /** 导出：按 DPI 渲染（PDF 以 72 为基准，Office 文档以 96 为屏幕基准）。 */
    fun renderPageDpi(index: Int, dpi: Int): Bitmap

    /**
     * 导出 PDF。
     * PDF 源走原生抽页（保矢量、文字、链接）；Office 源没有矢量可抽，
     * 按 WPS 的做法——每页先画成位图再用 PdfDocument 写页，视觉即所见。
     */
    fun exportPdf(selected: IntArray, dpi: Int, out: OutputStream) {
        require(selected.isNotEmpty()) { "没有选中任何页面" }
        val doc = PdfDocument()
        try {
            selected.forEachIndexed { i, index ->
                val bmp = renderPageDpi(index, dpi)
                try {
                    // PDF 单页上限 14400 单位：超高长表整体缩档，避免写出损坏文件
                    var w = bmp.width
                    var h = bmp.height
                    val over = maxOf(w, h) / MAX_PDF_SIDE.toFloat()
                    if (over > 1f) {
                        w = (w / over).toInt().coerceAtLeast(1)
                        h = (h / over).toInt().coerceAtLeast(1)
                    }
                    val page = doc.startPage(PdfDocument.PageInfo.Builder(w, h, i + 1).create())
                    page.canvas.drawColor(Color.WHITE)
                    page.canvas.drawBitmap(
                        bmp, null,
                        android.graphics.Rect(0, 0, w, h),
                        android.graphics.Paint(android.graphics.Paint.FILTER_BITMAP_FLAG)
                    )
                    doc.finishPage(page)
                } finally {
                    bmp.recycle()
                }
            }
            doc.writeTo(out)
        } finally {
            doc.close()
        }
    }

    companion object {
        private const val MAX_PDF_SIDE = 14000
    }
}

/** 打开失败时带一句能直接给用户看的话。 */
class PageSourceException(message: String) : Exception(message)

/**
 * 按扩展名分发。**先看扩展名再看 MIME**：SAF 里同一个文件在不同 ROM 上报的
 * MIME 可能千奇百怪（不少文件管理器统一报 application/octet-stream）。
 */
object PageSources {

    val SUPPORTED_EXT = arrayOf("pdf", "pptx", "xlsx", "xls")

    /** 文件选择器要过滤的类型：Office 旧格式也列出来，好给"请先另存为新格式"的提示。 */
    val MIME_TYPES = arrayOf(
        "application/pdf",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-powerpoint",
        "application/vnd.ms-excel",
        "application/octet-stream",
    )

    fun extOf(name: String): String = name.substringAfterLast('.', "").lowercase()

    fun open(context: Context, uri: Uri, name: String): PageSource {
        return when (extOf(name)) {
            "pdf" -> PdfDoc(context, uri)
            "pptx" -> PptxDoc(context, uri)
            "xlsx" -> XlsxDoc(context, uri)
            "xls" -> XlsDoc(context, uri)
            "ppt" -> throw PageSourceException("旧版 .ppt 请先用 Office/WPS 另存为 .pptx")
            "doc", "docx" -> throw PageSourceException("Android 版暂不支持 Word（分页要靠排版引擎算）")
            else -> throw PageSourceException("不支持的文件类型 .${extOf(name)}")
        }
    }
}
