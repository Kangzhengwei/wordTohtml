package com.example.exmword;

import android.app.Activity;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.view.View.OnClickListener;
import android.webkit.WebView;
import android.widget.TextView;

import com.kzw.wordtohtml.util.FileUtil;
import com.kzw.wordtohtml.word.WordUtil;

public class TextActivity extends Activity implements OnClickListener {
	private final static String TAG = "TextActivity";
	private WebView tv_content;
	private String documentPath = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS).getAbsolutePath() + "/dhms";

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_text);
		
		findViewById(R.id.btn_open).setOnClickListener(this);
		tv_content = (WebView) findViewById(R.id.wv_content);
	}

	@Override
	public void onClick(View v) {
		if (v.getId() == R.id.btn_open) {
			String filePath = documentPath + "/t3.doc";
			String htmlPath = FileUtil.createFile(documentPath, FileUtil.getFileName(filePath) + ".html");
			new WordUtil(filePath, htmlPath);
			Log.d(TAG, "htmlPath=" + htmlPath);
			tv_content.loadUrl("file:///" + htmlPath);
		}
	}

}
