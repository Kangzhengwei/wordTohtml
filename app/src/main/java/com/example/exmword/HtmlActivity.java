package com.example.exmword;

import android.app.Activity;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.view.View.OnClickListener;
import android.webkit.WebView;

import androidx.core.app.ActivityCompat;

import com.kzw.wordtohtml.excel.ExcelToHtmlUtils;
import com.kzw.wordtohtml.util.FileUtil;
import com.kzw.wordtohtml.word.WordUtil;


public class HtmlActivity extends Activity implements OnClickListener {
    private final static String TAG = "HtmlActivity";
    private WebView wv_content;
    private String documentPath = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS).getAbsolutePath() + "/dhms";

    private static final String[] PERMISSIONS_STORAGE = {"android.permission.READ_EXTERNAL_STORAGE",
            "android.permission.WRITE_EXTERNAL_STORAGE"};

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_html);
        ActivityCompat.requestPermissions(this, PERMISSIONS_STORAGE, 10);
        findViewById(R.id.btn_open).setOnClickListener(this);
        wv_content = (WebView) findViewById(R.id.wv_content);
    }

    @Override
    public void onClick(View v) {
        if (v.getId() == R.id.btn_open) {
            String filePath = documentPath + "/test1.docx";
            String htmlPath = FileUtil.createFile(documentPath, FileUtil.getFileName(filePath) + ".html");
            new WordUtil(filePath, htmlPath);
            Log.d(TAG, "htmlPath=" + htmlPath);
            wv_content.loadUrl("file:///" + htmlPath);
            /////////////////////////////////

        }
    }
}
