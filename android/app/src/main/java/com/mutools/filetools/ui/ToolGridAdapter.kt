package com.mutools.filetools.ui

import android.content.res.ColorStateList
import android.view.LayoutInflater
import android.view.ViewGroup
import androidx.annotation.ColorRes
import androidx.annotation.DrawableRes
import androidx.core.content.ContextCompat
import androidx.recyclerview.widget.RecyclerView
import com.mutools.filetools.databinding.ItemToolBinding

/** 一个工具的定义：id / 名称 / 分类 / 图标 / 底色 / 打开动作（工具注册制）。 */
data class ToolDef(
    val id: String,
    val label: String,
    val category: String,
    @DrawableRes val icon: Int,
    @ColorRes val tint: Int,
    val open: () -> Unit,
)

/** 首页工具网格：4 列图标 + 名称，支持按分类过滤（对应截图里的「全部工具」页）。 */
class ToolGridAdapter(
    private val tools: List<ToolDef>,
    private val onClick: (ToolDef) -> Unit,
) : RecyclerView.Adapter<ToolGridAdapter.VH>() {

    private var shown: List<ToolDef> = tools

    /** category=null 显示全部。 */
    fun filter(category: String?) {
        shown = if (category == null) tools else tools.filter { it.category == category }
        notifyDataSetChanged()
    }

    class VH(val binding: ItemToolBinding) : RecyclerView.ViewHolder(binding.root)

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): VH =
        VH(ItemToolBinding.inflate(LayoutInflater.from(parent.context), parent, false))

    override fun getItemCount(): Int = shown.size

    override fun onBindViewHolder(h: VH, pos: Int) {
        val def = shown[pos]
        h.binding.toolLabel.text = def.label
        h.binding.toolIcon.setImageResource(def.icon)
        h.binding.toolIcon.backgroundTintList =
            ColorStateList.valueOf(ContextCompat.getColor(h.binding.root.context, def.tint))
        h.binding.root.setOnClickListener { onClick(def) }
    }
}
