package com.mutools.filetools

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.provider.MediaStore
import android.view.View
import android.widget.Toast
import androidx.activity.OnBackPressedCallback
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.appcompat.app.AppCompatDelegate
import androidx.core.view.ViewCompat
import androidx.core.view.WindowCompat
import androidx.core.view.WindowInsetsCompat
import androidx.core.view.updatePadding
import androidx.recyclerview.widget.GridLayoutManager
import com.google.android.material.tabs.TabLayout
import com.mutools.filetools.databinding.ActivityMainBinding
import com.mutools.filetools.ui.ExtractPage
import com.mutools.filetools.ui.IcoPage
import com.mutools.filetools.ui.ToolDef
import com.mutools.filetools.ui.ToolGridAdapter
import com.mutools.filetools.util.Prefs

/**
 * 主壳（工具箱风格，对应参考截图的「全部工具」页）：
 * 居中标题栏 + 分类标签 + 工具网格 → 点卡片进工具子页（带返回栏）。
 * 工具注册制：tools 列表里登记即可新增工具（后续 Word/Excel/PPT 引擎在此接入）。
 */
class MainActivity : AppCompatActivity(), Host {

    private lateinit var binding: ActivityMainBinding
    private var extractPage: ExtractPage? = null
    private var icoPage: IcoPage? = null
    private lateinit var toolAdapter: ToolGridAdapter

    private val tools: List<ToolDef> = listOf(
        ToolDef("extract", "页面提取", "文档处理", R.drawable.ic_tool_extract, R.color.tile_blue) {
            showTool("extract", "页面提取")
        },
        ToolDef("ico", "图片转 ICO", "图片处理", R.drawable.ic_tool_ico, R.color.tile_orange) {
            showTool("ico", "图片转 ICO")
        },
    )

    // ------------------------------------------------------------ 选择器路由

    private val pdfPicker = registerForActivityResult(ActivityResultContracts.OpenMultipleDocuments()) { uris ->
        if (!uris.isNullOrEmpty()) extractPage?.onDocsAdded(uris)
    }
    private val imagePicker = registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri ->
        uri?.let { icoPage?.onImagePicked(it) }
    }
    private val galleryPicker =
        registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
            result.data?.data?.let { uri -> icoPage?.onImagePicked(uri) }
        }
    private val treePicker = registerForActivityResult(ActivityResultContracts.OpenDocumentTree()) { uri ->
        if (uri != null) {
            try {
                contentResolver.takePersistableUriPermission(
                    uri,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                )
            } catch (_: Exception) {
            }
            Prefs.setOutTree(this, uri.toString())
            extractPage?.onTreeChanged()
        }
    }

    private var createCallback: ((Uri?) -> Unit)? = null
    private val createPdf = registerForActivityResult(ActivityResultContracts.CreateDocument("application/pdf")) { u ->
        createCallback?.invoke(u); createCallback = null
    }
    private val createPng = registerForActivityResult(ActivityResultContracts.CreateDocument("image/png")) { u ->
        createCallback?.invoke(u); createCallback = null
    }
    private val createJpg = registerForActivityResult(ActivityResultContracts.CreateDocument("image/jpeg")) { u ->
        createCallback?.invoke(u); createCallback = null
    }
    private val createIco = registerForActivityResult(ActivityResultContracts.CreateDocument("image/x-icon")) { u ->
        createCallback?.invoke(u); createCallback = null
    }

    override fun pickDocs() = pdfPicker.launch(com.mutools.filetools.core.PageSources.MIME_TYPES)

    /** 相册优先：DocumentsUI 在很多国产 ROM 上看不到相册图片。 */
    override fun pickImage() {
        galleryPicker.launch(Intent(Intent.ACTION_PICK, MediaStore.Images.Media.EXTERNAL_CONTENT_URI))
    }

    /** 文件选择器兜底（相册之外/其他应用的图片）。 */
    override fun pickImageFile() = imagePicker.launch(arrayOf("image/*"))

    override fun pickTree() = treePicker.launch(null)

    override fun launchCreate(mime: String, suggestedName: String, onCreated: (Uri?) -> Unit) {
        createCallback = onCreated
        when (mime) {
            "application/pdf" -> createPdf.launch(suggestedName)
            "image/png" -> createPng.launch(suggestedName)
            "image/jpeg" -> createJpg.launch(suggestedName)
            else -> createIco.launch(suggestedName)
        }
    }

    // ------------------------------------------------------------ 生命周期

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        // Android 15 对 targetSdk 35 强制 edge-to-edge：自己避让系统栏，
        // 否则标题栏会顶进状态栏、底部导航会被手势条压住
        WindowCompat.setDecorFitsSystemWindows(window, false)
        ViewCompat.setOnApplyWindowInsetsListener(binding.root) { v, insets ->
            val bars = insets.getInsets(WindowInsetsCompat.Type.systemBars())
            val ime = insets.getInsets(WindowInsetsCompat.Type.ime())
            v.updatePadding(top = bars.top, bottom = maxOf(bars.bottom, ime.bottom))
            insets
        }

        extractPage = ExtractPage(this, this, binding.toolContainer)
        icoPage = IcoPage(this, this, binding.toolContainer)

        toolAdapter = ToolGridAdapter(tools) { def -> showTool(def.id, def.label) }
        binding.gridTools.layoutManager = GridLayoutManager(this, 4)
        binding.gridTools.adapter = toolAdapter

        val tabs = binding.tabs
        tabs.addTab(tabs.newTab().setText("全部"))
        tabs.addTab(tabs.newTab().setText("图片处理"))
        tabs.addTab(tabs.newTab().setText("文档处理"))
        tabs.addOnTabSelectedListener(object : TabLayout.OnTabSelectedListener {
            override fun onTabSelected(tab: TabLayout.Tab) {
                val cat = when (tab.position) {
                    1 -> "图片处理"
                    2 -> "文档处理"
                    else -> null
                }
                toolAdapter.filter(cat)
                binding.txtSection.text = cat ?: getString(R.string.section_common)
            }

            override fun onTabUnselected(tab: TabLayout.Tab) {}
            override fun onTabReselected(tab: TabLayout.Tab) {}
        })

        binding.btnBack.setOnClickListener { showHome() }
        binding.btnTheme.setOnClickListener { cycleTheme() }
        applyTheme(Prefs.themeMode(this))

        onBackPressedDispatcher.addCallback(this, object : OnBackPressedCallback(true) {
            override fun handleOnBackPressed() {
                if (binding.toolLayer.visibility == View.VISIBLE) showHome() else finish()
            }
        })

        if (savedInstanceState != null) {
            extractPage?.restore(savedInstanceState)
            icoPage?.restore(savedInstanceState)
            if (savedInstanceState.getBoolean("tool_open")) {
                val id = savedInstanceState.getString("tool_id") ?: "extract"
                tools.firstOrNull { it.id == id }?.let { showTool(it.id, it.label) }
            }
        }
    }

    override fun onSaveInstanceState(outState: Bundle) {
        super.onSaveInstanceState(outState)
        extractPage?.save(outState)
        icoPage?.save(outState)
        val open = binding.toolLayer.visibility == View.VISIBLE
        outState.putBoolean("tool_open", open)
        if (open) {
            outState.putString("tool_id", if (extractPage?.binding?.root?.visibility == View.VISIBLE) "extract" else "ico")
            outState.putString("tool_title", binding.txtToolTitle.text.toString())
        }
    }

    override fun onDestroy() {
        extractPage?.close()
        icoPage?.close()
        super.onDestroy()
    }

    // ------------------------------------------------------------ 首页 ↔ 工具页

    private fun showTool(id: String, label: String) {
        binding.homeScroll.visibility = View.GONE
        binding.tabs.visibility = View.GONE
        binding.toolLayer.visibility = View.VISIBLE
        binding.txtToolTitle.text = label
        extractPage?.binding?.root?.visibility = if (id == "extract") View.VISIBLE else View.GONE
        icoPage?.binding?.root?.visibility = if (id == "ico") View.VISIBLE else View.GONE
    }

    private fun showHome() {
        binding.toolLayer.visibility = View.GONE
        binding.homeScroll.visibility = View.VISIBLE
        binding.tabs.visibility = View.VISIBLE
    }

    // ------------------------------------------------------------ 主题

    private fun applyTheme(mode: String) {
        AppCompatDelegate.setDefaultNightMode(
            when (mode) {
                "light" -> AppCompatDelegate.MODE_NIGHT_NO
                "dark" -> AppCompatDelegate.MODE_NIGHT_YES
                else -> AppCompatDelegate.MODE_NIGHT_FOLLOW_SYSTEM
            }
        )
        binding.btnTheme.text = getString(
            when (mode) {
                "light" -> R.string.theme_button_light
                "dark" -> R.string.theme_button_dark
                else -> R.string.theme_button_system
            }
        )
    }

    private fun cycleTheme() {
        val next = when (Prefs.themeMode(this)) {
            "system" -> "light"
            "light" -> "dark"
            else -> "system"
        }
        Prefs.setThemeMode(this, next)
        applyTheme(next)
        val label = when (next) {
            "light" -> "浅色"
            "dark" -> "深色"
            else -> "跟随系统"
        }
        Toast.makeText(this, "外观：$label", Toast.LENGTH_SHORT).show()
    }
}

/** 工具页需要的宿主能力（选择器/建文件），MainActivity 实现。 */
interface Host {
    fun pickDocs()
    fun pickImage()
    fun pickImageFile()
    fun pickTree()
    fun launchCreate(mime: String, suggestedName: String, onCreated: (Uri?) -> Unit)
}
