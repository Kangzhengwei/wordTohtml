package com.example.exmword

import android.app.Activity
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.webkit.WebView
import com.kzw.wordtohtml.excel.ExcelToHtmlUtils

class ExcelxlsxActivity : Activity() {
    private val documentPath = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS).absolutePath + "/dhms"

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_excelxlsx)
        val vebView = findViewById<WebView>(R.id.webview)
        val path = ExcelToHtmlUtils.xlsxToHtml(documentPath, "test")
        vebView.loadUrl("file:///$documentPath/$path")
    }
}