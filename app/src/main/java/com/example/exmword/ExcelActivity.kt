package com.example.exmword

import android.annotation.SuppressLint
import android.app.Activity
import android.os.Bundle
import android.os.Environment
import android.webkit.WebView
import com.kzw.wordtohtml.excel.ExcelToHtmlUtils

class ExcelActivity : Activity() {
    private val documentPath = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS).absolutePath + "/dhms"

    @SuppressLint("MissingInflatedId")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_excel)
        val vebView = findViewById<WebView>(R.id.webview)
        val tableValueModel = ExcelToHtmlUtils.xlsToHtml(documentPath, "t4")
        vebView.loadUrl("file:///$documentPath/$tableValueModel")
    }

}