package com.mutools.filetools.ui

import android.graphics.Bitmap
import android.net.Uri
import android.os.Bundle
import android.provider.OpenableColumns
import android.view.View
import android.view.ViewGroup
import android.widget.ArrayAdapter
import android.widget.FrameLayout
import android.widget.ImageView
import android.widget.Toast
import androidx.appcompat.app.AlertDialog
import androidx.documentfile.provider.DocumentFile
import androidx.lifecycle.lifecycleScope
import androidx.recyclerview.widget.GridLayoutManager
import com.mutools.filetools.Host
import com.mutools.filetools.core.ImageOps
import com.mutools.filetools.core.Outputs
import com.mutools.filetools.core.PdfDoc
import com.mutools.filetools.databinding.PageExtractBinding
import com.mutools.filetools.util.Prefs
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.io.ByteArrayOutputStream
import java.util.concurrent.Executors

/**
 * 「页面提取」工具页（Android 版仅支持 PDF）：
 * 添加 PDF → 缩略图勾选 → 导出 PNG/JPG/长图/提取 PDF。
 * 三栏桌面布局在手机上折叠为单列：文件行 → 网格 → 导出设置卡。
 */
class ExtractPage(
    private val activity: androidx.appcompat.app.AppCompatActivity,
    private val host: Host,
    container: ViewGroup,
) {
    class Entry(val doc: PdfDoc, val uri: Uri, val name: String, val selected: BooleanArray)

    val binding = PageExtractBinding.inflate(activity.layoutInflater)
    private val entries = ArrayList<Entry>()
    private var cur = 0
    private var suppressAll = false
    private var exporting = false

    private val thumbExecutor = Executors.newSingleThreadExecutor()
    private val DPI_VALUES = intArrayOf(100, 150, 200, 300, 600)

    private val gridAdapter = PageAdapter(
        render = { entryId, page, cb ->
            thumbExecutor.execute {
                val e = synchronized(entries) { entries.getOrNull(entryId) }
                val bmp = try { e?.doc?.renderPage(page, 300) } catch (_: Exception) { null }
                activity.runOnUiThread { cb(bmp) }
            }
        },
        toggle = { p ->
            val e = entries.getOrNull(cur)
            if (e != null && p in e.selected.indices) {
                e.selected[p] = !e.selected[p]
                refreshCounts()
                e.selected[p]
            } else false
        },
        preview = { p -> showPreview(p) },
    )

    init {
        container.addView(binding.root)
        setup()
    }

    private fun toast(msg: String) = Toast.makeText(activity, msg, Toast.LENGTH_SHORT).show()

    private fun setup() {
        binding.btnAddPdf.setOnClickListener { host.pickPdfs() }
        binding.btnPickDir.setOnClickListener { host.pickTree() }
        binding.btnClear.setOnClickListener {
            if (entries.isEmpty()) return@setOnClickListener
            synchronized(entries) { entries.forEach { runCatching { it.doc.close() } }; entries.clear() }
            refreshFileSpinner()
        }

        binding.spinFiles.setOnItemSelectedListener(object : android.widget.AdapterView.OnItemSelectedListener {
            override fun onItemSelected(parent: android.widget.AdapterView<*>?, view: View?, pos: Int, id: Long) {
                cur = pos
                bindGrid()
            }

            override fun onNothingSelected(parent: android.widget.AdapterView<*>?) {}
        })

        val formats = activity.resources.getStringArray(com.mutools.filetools.R.array.formats)
        binding.spinFormat.adapter = ArrayAdapter(activity, android.R.layout.simple_spinner_item, formats)
            .apply { setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item) }
        binding.spinFormat.onItemSelectedListener = object : android.widget.AdapterView.OnItemSelectedListener {
            override fun onItemSelected(parent: android.widget.AdapterView<*>?, view: View?, pos: Int, id: Long) {
                binding.rowDpi.visibility = if (pos == 4) View.GONE else View.VISIBLE
            }

            override fun onNothingSelected(parent: android.widget.AdapterView<*>?) {}
        }

        val dpis = activity.resources.getStringArray(com.mutools.filetools.R.array.dpis)
        binding.spinDpi.adapter = ArrayAdapter(activity, android.R.layout.simple_spinner_item, dpis)
            .apply { setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item) }
        binding.spinDpi.setSelection(1)

        binding.chkAll.setOnCheckedChangeListener { _, c ->
            if (suppressAll) return@setOnCheckedChangeListener
            entries.getOrNull(cur)?.let { e ->
                e.selected.fill(c)
                gridAdapter.updateChecked { p -> e.selected.getOrNull(p) ?: false }
                refreshCounts()
            }
        }

        binding.gridPages.layoutManager = GridLayoutManager(activity, 3)
        binding.gridPages.adapter = gridAdapter

        binding.btnExport.setOnClickListener { startExport() }
        onTreeChanged()
        refreshFileSpinner()
    }

    // ------------------------------------------------------------ 文件管理

    fun onPdfsAdded(uris: List<Uri>) {
        activity.lifecycleScope.launch(Dispatchers.IO) {
            val opened = ArrayList<Entry>()
            for (u in uris) {
                try {
                    activity.contentResolver.takePersistableUriPermission(
                        u,
                        android.content.Intent.FLAG_GRANT_READ_URI_PERMISSION
                    )
                } catch (_: Exception) {
                }
                try {
                    val doc = PdfDoc(activity, u)
                    val n = doc.pageCount
                    opened += Entry(doc, u, nameOf(u), BooleanArray(n) { n <= 9 })
                } catch (e: Exception) {
                    withContext(Dispatchers.Main) { toast("无法打开：${nameOf(u)}（${e.javaClass.simpleName}）") }
                }
            }
            withContext(Dispatchers.Main) {
                synchronized(entries) { entries += opened }
                refreshFileSpinner()
                if (opened.isNotEmpty()) {
                    // 选中第一个新加入的文件
                    binding.spinFiles.setSelection(entries.size - opened.size)
                }
            }
        }
    }

    private fun nameOf(uri: Uri): String {
        var n: String? = null
        try {
            activity.contentResolver.query(uri, arrayOf(OpenableColumns.DISPLAY_NAME), null, null, null)
                ?.use { c -> if (c.moveToFirst()) n = c.getString(0) }
        } catch (_: Exception) {
        }
        return n ?: uri.lastPathSegment?.substringAfterLast('/') ?: "文档"
    }

    private fun refreshFileSpinner() {
        val names = synchronized(entries) { entries.map { it.name } }
        binding.spinFiles.adapter = ArrayAdapter(activity, android.R.layout.simple_spinner_item, names)
            .apply { setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item) }
        val has = names.isNotEmpty()
        binding.txtEmpty.visibility = if (has) View.GONE else View.VISIBLE
        binding.gridPages.visibility = if (has) View.VISIBLE else View.GONE
        refreshCounts()
        bindGrid()
    }

    private fun bindGrid() {
        val e = entries.getOrNull(cur)
        if (e == null) {
            gridAdapter.setEntry(-1, 0, { false })
        } else {
            gridAdapter.setEntry(cur, e.doc.pageCount) { p -> e.selected.getOrNull(p) ?: false }
        }
        refreshCounts()
    }

    private fun refreshCounts() {
        val e = entries.getOrNull(cur)
        suppressAll = true
        if (e == null) {
            binding.txtSelCount.text = ""
            binding.chkAll.isChecked = false
            binding.chkAll.isEnabled = false
        } else {
            val sel = e.selected.count { it }
            binding.txtSelCount.text = "已选 $sel / ${e.selected.size} 页"
            binding.chkAll.isEnabled = true
            binding.chkAll.isChecked = sel == e.selected.size && sel > 0
        }
        suppressAll = false
    }

    private fun showPreview(page: Int) {
        val e = entries.getOrNull(cur) ?: return
        activity.lifecycleScope.launch(Dispatchers.IO) {
            val bmp = try { e.doc.renderPage(page, 1080) } catch (_: Exception) { null }
            withContext(Dispatchers.Main) {
                if (bmp != null && !activity.isFinishing) {
                    val frame = FrameLayout(activity)
                    val iv = ImageView(activity)
                    iv.setImageBitmap(bmp)
                    iv.adjustViewBounds = true
                    val pad = (10 * resources_density()).toInt()
                    frame.setPadding(pad, pad, pad, pad)
                    frame.addView(
                        iv,
                        FrameLayout.LayoutParams(
                            FrameLayout.LayoutParams.MATCH_PARENT,
                            FrameLayout.LayoutParams.WRAP_CONTENT
                        )
                    )
                    AlertDialog.Builder(activity)
                        .setTitle("第 ${page + 1} 页 · ${e.name}")
                        .setView(frame)
                        .setPositiveButton("关闭", null)
                        .show()
                }
            }
        }
    }

    private fun resources_density() = activity.resources.displayMetrics.density

    // ------------------------------------------------------------ 导出

    fun onTreeChanged() {
        val t = Prefs.outTree(activity)
        val name = t?.let {
            try {
                DocumentFile.fromTreeUri(activity, Uri.parse(it))?.name
            } catch (_: Exception) {
                null
            }
        }
        binding.txtOutDir.text = name ?: getString(com.mutools.filetools.R.string.not_chosen)
    }

    private fun getString(res: Int) = activity.getString(res)

    private fun beginJob() {
        exporting = true
        activity.runOnUiThread {
            binding.btnExport.isEnabled = false
            binding.progressExport.visibility = View.VISIBLE
            binding.txtStatus.text = "准备导出…"
        }
    }

    private fun progress(done: Int, total: Int, msg: String) {
        activity.runOnUiThread {
            binding.progressExport.max = maxOf(total, 1)
            binding.progressExport.setProgressCompat(done, true)
            binding.txtStatus.text = "$done/$total · $msg"
        }
    }

    private fun finishJob(msg: String) {
        exporting = false
        activity.runOnUiThread {
            binding.progressExport.visibility = View.GONE
            binding.txtStatus.text = msg
            binding.btnExport.isEnabled = true
            Toast.makeText(activity, msg, Toast.LENGTH_LONG).show()
        }
    }

    private fun startExport() {
        if (exporting) return
        val selEntries = synchronized(entries) { entries.filter { it.selected.any { b -> b } } }
        if (selEntries.isEmpty()) {
            toast("请先勾选要导出的页面")
            return
        }
        val fmt = binding.spinFormat.selectedItemPosition
        val dpi = DPI_VALUES[binding.spinDpi.selectedItemPosition.coerceIn(0, DPI_VALUES.size - 1)]
        val prefix = Outputs.clean(binding.editPrefix.text.toString())
        val suffix = Outputs.clean(binding.editSuffix.text.toString())
        val tree = Prefs.outTree(activity)?.let {
            try {
                DocumentFile.fromTreeUri(activity, Uri.parse(it))
            } catch (_: Exception) {
                null
            }
        }

        when (fmt) {
            0, 1 -> {
                if (tree == null) {
                    toast("请先选择输出目录")
                    return
                }
                val png = fmt == 0
                val ext = if (png) "png" else "jpg"
                val total = selEntries.sumOf { it.selected.count { b -> b } }
                beginJob()
                activity.lifecycleScope.launch(Dispatchers.IO) {
                    var files = 0
                    var done = 0
                    try {
                        for (en in selEntries) {
                            val stem = Outputs.stem(Outputs.clean(en.name))
                            for (p in 0 until en.doc.pageCount) {
                                if (!en.selected[p]) continue
                                progress(done, total, "导出 ${en.name} · 第 ${p + 1} 页")
                                val bmp = try { en.doc.renderPageDpi(p, dpi) } catch (_: Exception) { null }
                                if (bmp != null) {
                                    val target = Outputs.uniqueIn(tree, "${prefix}${stem}_${p + 1}${suffix}.$ext")
                                    if (target != null && Outputs.writeBitmap(activity, target.uri, bmp, png)) files++
                                    bmp.recycle()
                                }
                                done++
                            }
                        }
                        finishJob("已导出 $files 个文件")
                    } catch (e: Exception) {
                        finishJob("导出失败：${e.message}")
                    }
                }
            }

            2, 3 -> {
                val png = fmt == 2
                val ext = if (png) "png" else "jpg"
                if (tree == null && selEntries.size == 1) {
                    val en = selEntries[0]
                    val suggested = "${prefix}${Outputs.stem(Outputs.clean(en.name))}_长图${suffix}.$ext"
                    host.launchCreate(Outputs.mimeOf(suggested), suggested) { uri ->
                        if (uri == null) return@launchCreate
                        exportLongImage(listOf(en to uri), png, ext, dpi)
                    }
                } else if (tree == null) {
                    toast("多个文件需要先选择输出目录")
                } else {
                    val pairs = selEntries.mapNotNull { en ->
                        val stem = Outputs.stem(Outputs.clean(en.name))
                        val target = Outputs.uniqueIn(tree, "${prefix}${stem}_长图${suffix}.$ext")
                        if (target == null) null else en to target.uri
                    }
                    exportLongImage(pairs, png, ext, dpi)
                }
            }

            else -> {
                if (tree == null && selEntries.size == 1) {
                    val en = selEntries[0]
                    val suggested = "${prefix}${Outputs.stem(Outputs.clean(en.name))}_提取${suffix}.pdf"
                    host.launchCreate("application/pdf", suggested) { uri ->
                        if (uri == null) return@launchCreate
                        exportPdf(listOf(en to uri))
                    }
                } else if (tree == null) {
                    toast("多个文件需要先选择输出目录")
                } else {
                    val pairs = selEntries.mapNotNull { en ->
                        val stem = Outputs.stem(Outputs.clean(en.name))
                        val target = Outputs.uniqueIn(tree, "${prefix}${stem}_提取${suffix}.pdf")
                        if (target == null) null else en to target.uri
                    }
                    exportPdf(pairs)
                }
            }
        }
    }

    private fun exportLongImage(pairs: List<Pair<Entry, Uri>>, png: Boolean, ext: String, dpi: Int) {
        if (pairs.isEmpty()) {
            toast("没有可导出的内容")
            return
        }
        val total = pairs.sumOf { (en, _) -> en.selected.count { b -> b } }
        beginJob()
        activity.lifecycleScope.launch(Dispatchers.IO) {
            var files = 0
            var done = 0
            try {
                for ((en, uri) in pairs) {
                    val pages = ArrayList<Bitmap>()
                    for (p in 0 until en.doc.pageCount) {
                        if (!en.selected[p]) continue
                        progress(done, total, "渲染 ${en.name} · 第 ${p + 1} 页")
                        try {
                            en.doc.renderPageDpi(p, dpi)?.let { pages += it }
                        } catch (_: Exception) {
                        }
                        done++
                    }
                    if (pages.isNotEmpty()) {
                        progress(done, total, "拼接 ${en.name} 长图")
                        val long = ImageOps.stitch(pages)
                        if (Outputs.writeBitmap(activity, uri, long, png)) files++
                        long.recycle()
                        pages.forEach { it.recycle() }
                    }
                }
                finishJob("已导出 $files 个文件")
            } catch (e: Exception) {
                finishJob("导出失败：${e.message}")
            }
        }
    }

    private fun exportPdf(pairs: List<Pair<Entry, Uri>>) {
        if (pairs.isEmpty()) {
            toast("没有可导出的内容")
            return
        }
        val total = pairs.sumOf { (en, _) -> en.selected.count { b -> b } }
        beginJob()
        activity.lifecycleScope.launch(Dispatchers.IO) {
            var files = 0
            var done = 0
            try {
                for ((en, uri) in pairs) {
                    progress(done, total, "提取 ${en.name}")
                    val bytes = ByteArrayOutputStream()
                    try {
                        val sel = IntArray(en.selected.size) { it }.filter { en.selected[it] }.toIntArray()
                        en.doc.extract(sel, bytes)
                        Outputs.writeBytes(activity, uri, bytes.toByteArray())
                        files++
                    } catch (_: Exception) {
                    }
                    done += en.selected.count { b -> b }
                    progress(done, total, "写出 ${en.name}")
                }
                finishJob("已导出 $files 个文件")
            } catch (e: Exception) {
                finishJob("导出失败：${e.message}")
            }
        }
    }

    // ------------------------------------------------------------ 状态保存

    fun save(state: Bundle) {
        val list = synchronized(entries) { entries.toList() }
        if (list.isEmpty()) return
        state.putStringArrayList("pdf_uris", ArrayList(list.map { it.uri.toString() }))
        state.putStringArrayList("pdf_sel", ArrayList(list.map { it.selected.joinToString("") { b -> if (b) "1" else "0" } }))
        state.putInt("pdf_cur", cur)
    }

    fun restore(state: Bundle) {
        val uris = state.getStringArrayList("pdf_uris") ?: return
        val sels = state.getStringArrayList("pdf_sel")
        val curIdx = state.getInt("pdf_cur", 0)
        activity.lifecycleScope.launch(Dispatchers.IO) {
            val reopened = ArrayList<Entry>()
            uris.forEachIndexed { i, s ->
                try {
                    val u = Uri.parse(s)
                    val doc = PdfDoc(activity, u)
                    val mask = sels?.getOrNull(i) ?: ""
                    reopened += Entry(doc, u, nameOf(u), BooleanArray(doc.pageCount) { mask.getOrNull(it) == '1' })
                } catch (_: Exception) {
                }
            }
            withContext(Dispatchers.Main) {
                synchronized(entries) { entries += reopened }
                refreshFileSpinner()
                if (curIdx in entries.indices) binding.spinFiles.setSelection(curIdx)
            }
        }
    }

    fun close() {
        thumbExecutor.shutdownNow()
        synchronized(entries) { entries.forEach { runCatching { it.doc.close() } } }
    }
}
