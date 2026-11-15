package com.mutools.filetools

import android.app.Application
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader

class App : Application() {
    override fun onCreate() {
        super.onCreate()
        // pdfbox-android 需要初始化才能在 Android 上使用字体资源等
        PDFBoxResourceLoader.init(applicationContext)
    }
}
