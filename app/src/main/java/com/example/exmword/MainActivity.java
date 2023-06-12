package com.example.exmword;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.view.View.OnClickListener;

public class MainActivity extends Activity implements OnClickListener {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        findViewById(R.id.btn_word_doc).setOnClickListener(this);
        findViewById(R.id.btn_word_docx).setOnClickListener(this);
        findViewById(R.id.btn_excel_xls).setOnClickListener(this);
        findViewById(R.id.btn_excel_xlsx).setOnClickListener(this);

    }

    @Override
    public void onClick(View v) {
        if (v.getId() == R.id.btn_word_doc) {
            Intent intent = new Intent(this, TextActivity.class);
            startActivity(intent);
        } else if (v.getId() == R.id.btn_word_docx) {
            Intent intent = new Intent(this, HtmlActivity.class);
            startActivity(intent);
        } else if (v.getId() == R.id.btn_excel_xls) {
            Intent intent = new Intent(this, ExcelActivity.class);
            startActivity(intent);
        } else if (v.getId() == R.id.btn_excel_xlsx) {
            Intent intent = new Intent(this, ExcelxlsxActivity.class);
            startActivity(intent);
        }
    }

}
