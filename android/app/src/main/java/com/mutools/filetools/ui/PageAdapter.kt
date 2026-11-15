package com.mutools.filetools.ui

import android.graphics.Bitmap
import android.util.LruCache
import android.view.LayoutInflater
import android.view.ViewGroup
import androidx.core.content.ContextCompat
import androidx.recyclerview.widget.RecyclerView
import com.mutools.filetools.R
import com.mutools.filetools.databinding.ItemPageBinding

/**
 * 页面缩略图网格：懒渲染（滚动到才渲染）、渲染结果按字节限量缓存、
 * 点按勾选、长按看大图。对应桌面版 ui/widgets.py 的单位卡片。
 */
class PageAdapter(
    private val render: (entryId: Int, page: Int, cb: (Bitmap?) -> Unit) -> Unit,
    private val toggle: (Int) -> Boolean,
    private val preview: (Int) -> Unit,
) : RecyclerView.Adapter<PageAdapter.VH>() {

    private var entryId = -1
    private var pageCount = 0
    private var checked: (Int) -> Boolean = { false }

    // 32MB 上限：300px 宽的缩略图约 0.5MB/张，最多留 60 张左右
    private val cache = object : LruCache<Int, Bitmap>(32 * 1024 * 1024) {
        override fun sizeOf(key: Int, value: Bitmap): Int = value.byteCount
    }

    fun setEntry(id: Int, pageCount: Int, checked: (Int) -> Boolean) {
        entryId = id
        this.pageCount = pageCount
        this.checked = checked
        cache.evictAll()
        notifyDataSetChanged()
    }

    fun updateChecked(checked: (Int) -> Boolean) {
        this.checked = checked
        notifyDataSetChanged()
    }

    class VH(val binding: ItemPageBinding) : RecyclerView.ViewHolder(binding.root)

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): VH =
        VH(ItemPageBinding.inflate(LayoutInflater.from(parent.context), parent, false))

    override fun getItemCount(): Int = pageCount

    override fun onBindViewHolder(h: VH, pos: Int) {
        val b = h.binding
        b.txtNo.text = "第 ${pos + 1} 页"
        applyCheck(b, checked(pos))
        b.root.setOnClickListener {
            val now = toggle(pos)
            if (h.bindingAdapterPosition != RecyclerView.NO_POSITION) applyCheck(b, now)
        }
        b.root.setOnLongClickListener { preview(pos); true }

        val key = entryId * 1_000_000 + pos
        val hit = cache.get(key)
        if (hit != null) {
            b.img.setImageBitmap(hit)
        } else {
            b.img.setImageDrawable(null)
            render(entryId, pos) { bmp ->
                if (bmp != null) cache.put(key, bmp)
                // 只有这个槽位还绑着同一页时才回填，避免复用后串页
                if (h.bindingAdapterPosition == pos && bmp != null) {
                    b.img.setImageBitmap(bmp)
                }
            }
        }
    }

    private fun applyCheck(b: ItemPageBinding, sel: Boolean) {
        b.chk.isChecked = sel
        if (sel) {
            b.root.strokeWidth = dp(b, 2)
            b.root.setStrokeColor(ContextCompat.getColor(b.root.context, R.color.accent))
            b.root.setCardBackgroundColor(ContextCompat.getColor(b.root.context, R.color.surface_sunk))
        } else {
            b.root.strokeWidth = dp(b, 1)
            b.root.setStrokeColor(ContextCompat.getColor(b.root.context, R.color.border))
            b.root.setCardBackgroundColor(ContextCompat.getColor(b.root.context, R.color.surface))
        }
    }

    private fun dp(b: ItemPageBinding, v: Int): Int =
        (v * b.root.context.resources.displayMetrics.density).toInt().coerceAtLeast(1)
}
