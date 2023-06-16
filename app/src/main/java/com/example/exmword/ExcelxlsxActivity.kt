package com.example.exmword

import android.app.Activity
import android.os.Bundle
import android.os.Environment
import android.webkit.WebView

class ExcelxlsxActivity : Activity() {
    private val documentPath = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS).absolutePath + "/dhms"

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_excelxlsx)
        val vebView = findViewById<WebView>(R.id.webview)
    }
}