package com.kzw.wordtohtml.util;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import android.os.Environment;
import android.util.Log;

public class FileUtil {
    private final static String TAG = "FileUtil";

    public static String getFileName(String filePath) {
        int start = filePath.lastIndexOf("/");
        int end = filePath.lastIndexOf(".");
        if (start != -1 && end != -1) {
            return filePath.substring(start + 1, end);
        } else {
            return "";
        }
    }

    public static String createFile(String dirPath, String fileName) {
        String filePath = String.format("%s/%s", dirPath, fileName);
        try {
            File dirFile = new File(dirPath);
            if (!dirFile.exists()) {
                dirFile.mkdir();
            }
            File myFile = new File(filePath);
            myFile.createNewFile();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return filePath;
    }

    public static ZipEntry getPicEntry(ZipFile docxFile, int pic_index) {
        String entry_jpeg = "word/media/image" + pic_index + ".jpeg";
        String entry_jpg = "word/media/image" + pic_index + ".jpg";
        String entry_png = "word/media/image" + pic_index + ".png";
        String entry_gif = "word/media/image" + pic_index + ".gif";
        String entry_wmf = "word/media/image" + pic_index + ".wmf";
        ZipEntry pic_entry;
        pic_entry = docxFile.getEntry(entry_jpeg);
        // 以下为读取docx的图片 转化为流数组
        if (pic_entry == null) {
            pic_entry = docxFile.getEntry(entry_jpg);
        }
        if (pic_entry == null) {
            pic_entry = docxFile.getEntry(entry_png);
        }
        if (pic_entry == null) {
            pic_entry = docxFile.getEntry(entry_gif);
        }
        if (pic_entry == null) {
            pic_entry = docxFile.getEntry(entry_wmf);
        }
        return pic_entry;
    }

    public static byte[] getPictureBytes(ZipFile docxFile, ZipEntry pic_entry) {
        byte[] pictureBytes = null;
        try {
            InputStream pictIS = docxFile.getInputStream(pic_entry);
            ByteArrayOutputStream pOut = new ByteArrayOutputStream();
            byte[] b = new byte[1000];
            int len;
            while ((len = pictIS.read(b)) != -1) {
                pOut.write(b, 0, len);
            }
            pictIS.close();
            pOut.close();
            pictureBytes = pOut.toByteArray();
            Log.d(TAG, "pictureBytes.length=" + pictureBytes.length);
            pictIS.close();
            pOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return pictureBytes;

    }

    public static void writePicture(String pic_path, byte[] pictureBytes) {
        File myPicture = new File(pic_path);
        try {
            FileOutputStream outputPicture = new FileOutputStream(myPicture);
            outputPicture.write(pictureBytes);
            outputPicture.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
