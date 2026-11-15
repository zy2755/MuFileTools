import java.util.Properties

plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
}

android {
    namespace = "com.mutools.filetools"
    compileSdk = 35

    defaultConfig {
        applicationId = "com.mutools.filetools"
        minSdk = 26
        targetSdk = 35
        versionCode = 6
        versionName = "1.2.0"
    }

    signingConfigs {
        create("release") {
            // 签名信息不写进仓库：优先读 android/local.properties（.gitignore 已排除），
            // 其次读环境变量 MUKEYSTORE_PASS（CI 用）。
            // 缺密钥库或口令只影响 assembleRelease，assembleDebug 照常可用（用 debug 签名）。
            val local = Properties().apply {
                val f = rootProject.file("local.properties")
                if (f.exists()) f.inputStream().use { load(it) }
            }
            val pass = local.getProperty("keystore.pass") ?: System.getenv("MUKEYSTORE_PASS")
            if (pass.isNullOrBlank()) {
                logger.warn(
                    "[签名] 未配置 keystore.pass —— 请在 android/local.properties 里写 " +
                        "keystore.pass=… ，或设置环境变量 MUKEYSTORE_PASS；assembleRelease 会因缺口令失败。"
                )
            }
            storeFile = file(local.getProperty("keystore.path") ?: "../keystore/mufiletools.keystore")
            storePassword = pass
            keyAlias = local.getProperty("keystore.alias") ?: "mufiletools"
            keyPassword = pass
        }
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            signingConfig = signingConfigs.getByName("release")
        }
    }

    buildFeatures {
        viewBinding = true
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    kotlinOptions {
        jvmTarget = "17"
    }
}

dependencies {
    implementation("androidx.core:core-ktx:1.13.1")
    implementation("androidx.appcompat:appcompat:1.7.0")
    implementation("com.google.android.material:material:1.12.0")
    implementation("androidx.constraintlayout:constraintlayout:2.1.4")
    implementation("androidx.recyclerview:recyclerview:1.3.2")
    implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.8.1")
    implementation("androidx.lifecycle:lifecycle-runtime-ktx:2.8.6")
    implementation("androidx.exifinterface:exifinterface:1.3.7")
    implementation("androidx.documentfile:documentfile:1.0.1")
    implementation("com.tom-roush:pdfbox-android:2.0.27.0")
    testImplementation("junit:junit:4.13.2")
}
