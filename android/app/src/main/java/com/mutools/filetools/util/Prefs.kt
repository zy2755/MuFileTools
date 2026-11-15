package com.mutools.filetools.util

import android.content.Context

/** 轻量偏好：外观模式（system/light/dark）、导出目录树 URI。 */
object Prefs {

    private const val NAME = "prefs"

    private fun sp(ctx: Context) = ctx.getSharedPreferences(NAME, Context.MODE_PRIVATE)

    fun themeMode(ctx: Context): String = sp(ctx).getString("theme_mode", "system") ?: "system"

    fun setThemeMode(ctx: Context, v: String) {
        sp(ctx).edit().putString("theme_mode", v).apply()
    }

    fun outTree(ctx: Context): String? = sp(ctx).getString("out_tree", null)

    fun setOutTree(ctx: Context, v: String?) {
        sp(ctx).edit().putString("out_tree", v).apply()
    }
}
