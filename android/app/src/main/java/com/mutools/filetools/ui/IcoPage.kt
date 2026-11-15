package com.mutools.filetools.ui

import android.graphics.Bitmap
import android.net.Uri
import android.os.Bundle
import android.provider.OpenableColumns
import android.view.ViewGroup
import android.widget.Toast
import androidx.lifecycle.lifecycleScope
import com.mutools.filetools.Host
import com.mutools.filetools.core.IcoCore
import com.mutools.filetools.core.IcoEncoder
import com.mutools.filetools.core.ImageOps
import com.mutools.filetools.core.Outputs
import com.mutools.filetools.databinding.IcoPageBinding
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.util.concurrent.Executors
import java.util.concurrent.atomic.AtomicBoolean

/**
 * 「图片转 ICO」工具页：圆角滑杆实时预览（棋盘格衬托透明区）、
 * 居中裁剪、导出 256~16 七档 .ico。对应桌面版 ui/ico_page.py。
 */
class IcoPage(
    private val activity: androidx.appcompat.app.AppCompatActivity,
    private val host: Host,
    container: ViewGroup,
) {
    val binding = IcoPageBinding.inflate(activity.layoutInflater)
    private var src: Bitmap? = null      // 已纠偏、≤1024
    private var square: Bitmap? = null   // 方形化后
    private var pct = 0
    private var sourceName: String? = null
    private var sourceUri: Uri? = null   // 主题切换重建后重新解码用

    private val exec = Executors.newSingleThreadExecutor()
    private val pending = AtomicBoolean(false)

    init {
        container.addView(binding.root)
        binding.btnPickImage.setOnClickListener { host.pickImage() }
        binding.btnPickFile.setOnClickListener { host.pickImageFile() }
        binding.sliderRound.addOnChangeListener { _, value, _ ->
            pct = value.toInt()
            refresh()
        }
        binding.chkCrop.setOnCheckedChangeListener { _, crop ->
            val s = src ?: return@setOnCheckedChangeListener
            square = IcoCore.toSquare(s, crop)
            refresh()
        }
        binding.btnExportIco.setOnClickListener { export() }
    }

    private fun toast(msg: String) = Toast.makeText(activity, msg, Toast.LENGTH_SHORT).show()

    fun onImagePicked(uri: Uri) {
        try {
            activity.contentResolver.takePersistableUriPermission(
                uri, android.content.Intent.FLAG_GRANT_READ_URI_PERMISSION
            )
        } catch (_: Exception) {
        }
        sourceUri = uri
        activity.lifecycleScope.launch(Dispatchers.IO) {
            var err: String? = null
            val b = try {
                ImageOps.decodeUpright(activity, uri, IcoCore.SOURCE_MAX)
            } catch (e: Exception) {
                android.util.Log.e("MuFileTools", "读取图片失败", e)
                err = "${e.javaClass.simpleName}: ${e.message ?: ""}"
                null
            }
            withContext(Dispatchers.Main) {
                if (b == null) {
                    toast("读取失败：${err ?: "格式不支持或图片为空"}，可换一张或用「从文件选择」重试")
                } else {
                    src?.recycle()
                    src = b
                    sourceName = nameOf(uri)
                    binding.txtImageName.text = sourceName
                    square = IcoCore.toSquare(b, binding.chkCrop.isChecked)
                    refresh()
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
        return n ?: uri.lastPathSegment?.substringAfterLast('/') ?: "图片"
    }

    /** 滑杆快速拖动时合并刷新：后到的计算覆盖先到的。 */
    private fun refresh() {
        val s = square ?: return
        if (!pending.compareAndSet(false, true)) return
        exec.execute {
            pending.set(false)
            val cur = pct
            val big = IcoCore.applyRound(s, cur.toFloat())
            val pv256 = IcoCore.onCheckerboard(IcoCore.applyRound(IcoCore.fitSquare(s, 256), cur.toFloat()))
            val pv64 = IcoCore.onCheckerboard(IcoCore.applyRound(IcoCore.fitSquare(s, 64), cur.toFloat()))
            val pv32 = IcoCore.onCheckerboard(IcoCore.applyRound(IcoCore.fitSquare(s, 32), cur.toFloat()))
            val pv16 = IcoCore.onCheckerboard(IcoCore.applyRound(IcoCore.fitSquare(s, 16), cur.toFloat()))
            activity.runOnUiThread {
                binding.imgMain.setImageBitmap(big)
                binding.pv256.setImageBitmap(pv256)
                binding.pv64.setImageBitmap(pv64)
                binding.pv32.setImageBitmap(pv32)
                binding.pv16.setImageBitmap(pv16)
                binding.txtPct.text = "$cur%"
            }
        }
    }

    private fun export() {
        val s = square ?: run {
            toast("请先选择图片")
            return
        }
        binding.btnExportIco.isEnabled = false
        exec.execute {
            val frames = IcoCore.buildFrames(s, pct.toFloat())
            val bytes = IcoEncoder.encode(frames)
            activity.runOnUiThread {
                binding.btnExportIco.isEnabled = true
                val stem = Outputs.clean(Outputs.stem(sourceName ?: "图标")).ifEmpty { "图标" }
                host.launchCreate(Outputs.mimeOf("$stem.ico"), "$stem.ico") { uri ->
                    if (uri == null) return@launchCreate
                    activity.lifecycleScope.launch(Dispatchers.IO) {
                        Outputs.writeBytes(activity, uri, bytes)
                        withContext(Dispatchers.Main) {
                            toast("已导出 ICO（七档尺寸）")
                        }
                    }
                }
            }
        }
    }

    fun save(state: Bundle) {
        sourceUri?.let { state.putString("ico_uri", it.toString()) }
        state.putInt("ico_pct", pct)
    }

    fun restore(state: Bundle) {
        pct = state.getInt("ico_pct", 0)
        val uri = state.getString("ico_uri")
        if (uri != null) {
            onImagePicked(Uri.parse(uri))
        } else {
            binding.sliderRound.value = pct.coerceIn(0, 50).toFloat()
        }
    }

    fun close() {
        exec.shutdownNow()
    }
}
