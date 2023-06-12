package com.kzw.wordtohtml.excel;

import android.os.Build;
import android.text.TextUtils;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ExcelToHtmlUtils {

    // 合并单元格后不生成<td>
    private static final String INVALID_CELL = "ExcelToHtml_invalidCell";
    private static final String CURLY_BRACES_LEFT = "{";
    private static final String BRACKETS_RIGHT_AND_CURLY_BRACES_LEFT = "){";
    private static final String CLASS_NAME_SHEET = ".sheet";
    private static final String WIDTH_COLON = "width:";
    private static final String CURLY_BRACES_RIGHT = "}";
    private static final String PX_SEMICOLON = "px;";
    /**
     * 字体样式属性名，带有:
     **/
    private static final String FONT_SIZE = "font-size:";
    /**
     * 样式class 选择器
     **/
    private static final String CLASS_SELECTOR = ".";
    /**
     * 背景颜色class前缀
     **/
    private static final String BACKGROUND_COLOR_CLASS_PREFIX = "bc";
    /**
     * 字体颜色class前缀
     **/
    private static final String COLOR_CLASS_PREFIX = "c";
    /**
     * 字体大小class前缀
     **/
    private static final String FONT_SIZE_CLASS_PREFIX = "fs";
    /**
     * 文本居中样式class
     **/
    private static final String TEXT_ALIGN_CENTER_CLASS = "tac";
    /**
     * 文本右对齐样式class
     **/
    private static final String TEXT_ALIGN_RIGHT_CLASS = "tar";
    /**
     * 字体加粗样式class
     **/
    private static final String BOLD_CLASS = "bold";
    /**
     * 斜体样式class
     **/
    private static final String ITALIC_CLASS = "italic";
    /**
     * 删除线样式class
     **/
    private static final String STRIKEOUT_CLASS = "strikeout";
    /**
     * 下划线样式class
     **/
    private static final String UNDERLINE_CLASS = "underline";
    private static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("0");

    /**
     * excel——xls转HTML
     *
     * @param excelPath excel文件所在路径（不包括文件名）
     * @param excelName excel文件名（不包括文件扩展名）
     * @return java.lang.String HTML储存路径
     * @author Kirito丶城
     * @date 2022/8/29
     */
    public static String xlsToHtml(String excelPath, String excelName) {
        String htmlPath = excelPath + File.separator + excelName + "_show" + File.separator;
        String htmlName = excelName + ".html";
        // 判断文件夹是否存在 不存在则创建文件夹
        File htmlFilePath = new File(htmlPath);
        if (!htmlFilePath.exists()) {
            htmlFilePath.mkdirs();
        }
        // 存储全部自定义样式内容。 key-class名  value-对应样式
        Map<String, String> cssMap = new HashMap<>();
        HSSFWorkbook hssfWorkbook = null;
        InputStream hssfStream = null;
        FileWriter cssFileWriter = null;
        BufferedWriter htmlBufferedWriter = null;
        try {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
                hssfStream = Files.newInputStream(Paths.get(excelPath, excelName + ".xls"));
            } else {
                hssfStream = new FileInputStream(new File(excelPath, excelName + ".xls"));
            }
            hssfWorkbook = new HSSFWorkbook(hssfStream);
            StringBuilder html = new StringBuilder("<!DOCTYPE html><html><head><meta charset='UTF-8'><title>");
            StringBuilder style = new StringBuilder();
            StringBuilder sheetTabHtml = new StringBuilder();
            html.append(excelName).append(".xls");
            html.append("</title>");
            html.append("<link rel='stylesheet' type='text/css' href='");
            html.append(excelPath);
            html.append("css/excel-sheet.css'/>");
            html.append("<link rel='stylesheet' type='text/css' href='excel-table.css'/>");
            html.append("<script src='");
            html.append(excelPath);
            html.append("js/plugins/jquery/jquery.min.js'></script>");
            html.append("<script src='");
            html.append(excelPath);
            html.append("js/excel-sheet.js'></script>");
            html.append("</head><body>");
            for (int st = 0; st < hssfWorkbook.getNumberOfSheets(); st++) {
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(st);
                String sheetName = hssfSheet.getSheetName();
                sheetTabHtml.append("<li ");
                sheetTabHtml.append(st == 0 ? "class='current'" : "");
                sheetTabHtml.append("sheet-filter='sheet");
                sheetTabHtml.append(st + 1);
                sheetTabHtml.append("'>");
                sheetTabHtml.append(sheetName);
                sheetTabHtml.append("</li>");
                List<List<Object>> tableData = new ArrayList<>();
                StringBuilder tableCssStyle = new StringBuilder();
                // 用于存储当前sheet中全部的单元格class
                List<List<List<String>>> tableRowCellClass = new ArrayList<>();
                // 储存各个列的宽度
                Map<Integer, Float> columnWidth = new HashMap<>();
                // 查找出每行最大列数量，防止出现每行列数量不同，导致表格无法渲染的bug
                short maxLastCellNum = 0;
                for (Row row : hssfSheet) {
                    short lastCellNum = row.getLastCellNum();
                    maxLastCellNum = lastCellNum > maxLastCellNum ? lastCellNum : maxLastCellNum;
                }
                for (Row row : hssfSheet) {
                    // 用于存储当前行的所有单元格的样式
                    List<List<String>> rowCellClass = new ArrayList<>();
                    // 初始化行数据数组
                    List<Object> rowValues = new ArrayList<>();
                    for (int i = 0; i < maxLastCellNum; i++) {
                        rowValues.add("");
                    }
                    for (Cell cell : row) {
                        // 用于存储当前单元格的全部class
                        List<String> cellClass = new ArrayList<>();
                        CellStyle cellStyle = cell.getCellStyle();

                        // 单元格宽度
                        int columnIndex = cell.getColumnIndex();
                        int columnWidthKey = columnIndex + 1;
                        float width = hssfSheet.getColumnWidthInPixels(columnIndex);
                        width = (width / 72) * 96;
                        Float colWidth = columnWidth.get(columnWidthKey);
                        if (colWidth == null || width > colWidth) {
                            columnWidth.put(columnWidthKey, width);
                        }

                        // 设置单元格样式
                        createCellStyle(cellClass, cellStyle);

                        // 填充颜色
                        if (cellStyle.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
                            Color color = cellStyle.getFillForegroundColorColor();
                            HSSFColor hssfColor = HSSFColor.toHSSFColor(color);
                            if (hssfColor != null) {
                                short[] triplet = hssfColor.getTriplet();
                                String className = COLOR_CLASS_PREFIX + triplet[0] + triplet[1] + triplet[2];
                                cssMap.put(className, "background-color:rgb(" + triplet[0] + "," + triplet[1] + "," + triplet[2] + ");");
                                cellClass.add(className);
                            }
                        }

                        HSSFFont hssfFont = hssfWorkbook.getFontAt(cellStyle.getFontIndex());

                        addFontStyle(cssMap, hssfFont, cellClass);

                        // 字体颜色
                        //  HSSFColor hssfColor = hssfFont.getHSSFColor(hssfWorkbook);
                        // if (hssfColor != null) {
                        // short[] triplet = hssfColor.getTriplet();
                           /* if (triplet[0] != 0 || triplet[1] != 0 || triplet[2] != 0) {
                                String className = COLOR_CLASS_PREFIX + triplet[0] + triplet[1] + triplet[2];
                                cssMap.put(className, "color:rgb(" + triplet[0] + "," + triplet[1] + "," + triplet[2] + ");");
                                cellClass.add(className);
                            }*/
                        //  }

                        setValues(rowValues, cell);
                        rowCellClass.add(cellClass);
                    }
                    tableData.add(rowValues);
                    tableRowCellClass.add(rowCellClass);
                }
                // 储存存在合并单元格的坐标  这些行已经转译玩成HTML
                List<String> mergedRegions = getMergedRegions(hssfSheet, tableData, tableRowCellClass);
                boolean hasMerged = CollectionUtils.isNotEmpty(mergedRegions);
                // 修改表格数据，清除仅有样式没有值的多余列和行
                tableData = trimTableData(tableData, columnWidth);
                // 添加列宽样式
                appendColWidthCss(tableCssStyle, st, columnWidth);

                style.append(tableCssStyle);
                html.append("<div class='sheet sheet");
                html.append(st + 1);
                html.append(st == 0 ? " sheet-show" : "");
                html.append("'>");
                // 生成表格HTML
                html.append(createTableHtml(tableData, tableRowCellClass, hasMerged, mergedRegions));
                html.append("</div>");
            }
            html.append("<div class='sheet-tab'><button class='roll-left' disabled='disabled'>◀</button><ul>");
            html.append(sheetTabHtml);
            html.append("</ul><button class='roll-right'>▶</button></div>");
            html.append("</body></html>");
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
                htmlBufferedWriter = new BufferedWriter(new OutputStreamWriter(Files.newOutputStream(Paths.get(htmlPath + htmlName)), StandardCharsets.UTF_8));
            } else {
                htmlBufferedWriter = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File(htmlPath + htmlName)), StandardCharsets.UTF_8));
            }
            htmlBufferedWriter.write(html.toString());

            // 补充字体样式
            appendCss(style, cssMap);
            cssFileWriter = new FileWriter(htmlPath + "excel-table.css");
            cssFileWriter.write(style.toString());
            return excelName + "_show/" + htmlName;
        } catch (IOException e) {

        } finally {
            if (hssfWorkbook != null) {
                try {
                    hssfWorkbook.close();
                } catch (IOException e) {

                }
            }
            if (hssfStream != null) {
                try {
                    hssfStream.close();
                } catch (IOException e) {

                }
            }
            closeFileWriter(htmlBufferedWriter, cssFileWriter);
        }
        return "";
    }

    /**
     * excel——xlsx转HTML
     *
     * @param excelPath excel文件所在路径（不包括文件名）
     * @param excelName excel文件名（不包括文件扩展名）
     * @return java.lang.String HTML保存路径
     * @author Kirito丶城
     * @date 2022/8/29
     */
    public static String xlsxToHtml(String excelPath, String excelName) {
        String htmlPath = excelPath + File.separator + excelName + "_show" + File.separator;
        String htmlName = excelName + ".html";
        // 判断文件夹是否存在 不存在则创建文件夹
        File htmlFilePath = new File(htmlPath);
        if (!htmlFilePath.exists()) {
            htmlFilePath.mkdirs();
        }
        // 存储全部自定义样式内容。 key-class名  value-对应样式
        Map<String, String> cssMap = new HashMap<>();
        InputStream xssfStream = null;
        XSSFWorkbook xssfWorkbook = null;
        FileWriter cssFileWriter = null;
        BufferedWriter htmlBufferedWriter = null;
        try {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
                xssfStream = Files.newInputStream(Paths.get(excelPath, excelName + ".xlsx"));
            } else {
                xssfStream = new FileInputStream(new File(excelPath, excelName + ".xlsx"));
            }
            xssfWorkbook = new XSSFWorkbook(xssfStream);
            StringBuilder html = new StringBuilder("<!DOCTYPE html><html><head><meta charset='UTF-8'><title>");
            StringBuilder style = new StringBuilder();
            html.append(excelName).append(".xlsx");
            html.append("</title>");
            html.append("<link rel='stylesheet' type='text/css' href='");
            html.append(excelPath);
            html.append("css/excel-sheet.css'/>");
            html.append("<link rel='stylesheet' type='text/css' href='excel-table.css'/>");
            html.append("<script src='");
            html.append(excelPath);
            html.append("js/plugins/jquery/jquery.min.js'></script>");
            html.append("<script src='");
            html.append(excelPath);
            html.append("js/excel-sheet.js'></script>");
            html.append("</head><body>");
            StringBuilder sheetTabHtml = new StringBuilder();
            for (int st = 0; st < xssfWorkbook.getNumberOfSheets(); st++) {
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(st);
                String sheetName = xssfSheet.getSheetName();
                sheetTabHtml.append("<li ");
                sheetTabHtml.append(st == 0 ? "class='current'" : "");
                sheetTabHtml.append("sheet-filter='sheet");
                sheetTabHtml.append(st + 1);
                sheetTabHtml.append("'>");
                sheetTabHtml.append(sheetName);
                sheetTabHtml.append("</li>");
                List<List<Object>> tableData = new ArrayList<>();
                StringBuilder tableCssStyle = new StringBuilder();
                // 用于存储当前sheet中全部的单元格class
                List<List<List<String>>> tableRowCellClass = new ArrayList<>();
                // 储存各个列的宽度
                Map<Integer, Float> columnWidth = new HashMap<>();
                // 查找出每行最大列数量，防止出现每行列数量不同，导致表格无法渲染的bug
                short maxLastCellNum = 0;
                for (Row row : xssfSheet) {
                    short lastCellNum = row.getLastCellNum();
                    maxLastCellNum = lastCellNum > maxLastCellNum ? lastCellNum : maxLastCellNum;
                }
                for (Row row : xssfSheet) {
                    // 用于存储当前行的所有单元格的样式
                    List<List<String>> rowCellClass = new ArrayList<>();
                    // 初始化行数据数组
                    List<Object> rowValues = new ArrayList<>();
                    for (int i = 0; i < maxLastCellNum; i++) {
                        rowValues.add("");
                    }
                    // 列编号
                    for (Cell cell : row) {
                        // 用于存储当前单元格的全部class
                        List<String> cellClass = new ArrayList<>();
                        CellStyle cellStyle = cell.getCellStyle();

                        // 单元格宽度
                        int columnIndex = cell.getColumnIndex();
                        int columnWidthKey = columnIndex + 1;
                        float width = xssfSheet.getColumnWidthInPixels(columnIndex);
                        width = (width / 72) * 96;
                        Float colWidth = columnWidth.get(columnWidthKey);
                        if (colWidth == null || width > colWidth) {
                            columnWidth.put(columnWidthKey, width);
                        }

                        // 设置单元格样式
                        createCellStyle(cellClass, cellStyle);

                        // 填充颜色
                        if (cellStyle.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
                            Color color = cellStyle.getFillForegroundColorColor();
                            XSSFColor xssfColor = XSSFColor.toXSSFColor(color);
                            if (xssfColor != null) {
                                String hex = xssfColor.getARGBHex().substring(2);
                                String className = BACKGROUND_COLOR_CLASS_PREFIX + hex;
                                cssMap.put(className, "background-color:#" + hex + ";");
                                cellClass.add(className);
                            }
                        }

                        XSSFFont xssfFont = xssfWorkbook.getFontAt(cellStyle.getFontIndex());

                        addFontStyle(cssMap, xssfFont, cellClass);

                        // 字体颜色
                        XSSFColor xssfColor = xssfFont.getXSSFColor();
                        if (xssfColor != null) {
                            String hex = xssfColor.getARGBHex().substring(2);
                            if (!"000000".equals(hex)) {
                                String className = COLOR_CLASS_PREFIX + hex;
                                cssMap.put(className, "color:#" + hex + ";");
                                cellClass.add(className);
                            }
                        }

                        setValues(rowValues, cell);
                        rowCellClass.add(cellClass);
                    }
                    tableData.add(rowValues);
                    tableRowCellClass.add(rowCellClass);
                }

                // 储存存在合并单元格的坐标  这些行已经转译玩成HTML
                List<String> mergedRegions = getMergedRegions(xssfSheet, tableData, tableRowCellClass);
                boolean hasMerged = CollectionUtils.isNotEmpty(mergedRegions);
                // 修改表格数据，清除仅有样式没有值的多余列和行
                tableData = trimTableData(tableData, columnWidth);
                // 添加列宽样式
                appendColWidthCss(tableCssStyle, st, columnWidth);

                style.append(tableCssStyle);
                html.append("<div class='sheet sheet");
                html.append(st + 1);
                html.append(st == 0 ? " sheet-show" : "");
                html.append("'>");
                // 生成表格HTML
                html.append(createTableHtml(tableData, tableRowCellClass, hasMerged, mergedRegions));
                html.append("<br>");
                html.append("</div>");
            }
            html.append("<div class='sheet-tab'><button class='roll-left' disabled='disabled'>◀</button><ul>");
            html.append(sheetTabHtml);
            html.append("</ul><button class='roll-right'>▶</button></div>");
            html.append("</body></html>");
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
                htmlBufferedWriter = new BufferedWriter(new OutputStreamWriter(Files.newOutputStream(Paths.get(htmlPath + htmlName)), StandardCharsets.UTF_8));
            } else {
                htmlBufferedWriter = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(htmlPath + htmlName), StandardCharsets.UTF_8));
            }
            htmlBufferedWriter.write(html.toString());
            // 补充字体样式
            appendCss(style, cssMap);
            cssFileWriter = new FileWriter(htmlPath + "excel-table.css");
            cssFileWriter.write(style.toString());
            return excelName + "_show/" + htmlName;
        } catch (IOException e) {

        } finally {
            if (xssfWorkbook != null) {
                try {
                    xssfWorkbook.close();
                } catch (IOException e) {

                }
            }
            if (xssfStream != null) {
                try {
                    xssfStream.close();
                } catch (IOException e) {

                }
            }
            closeFileWriter(htmlBufferedWriter, cssFileWriter);
        }
        return "";
    }

    /**
     * 用于生产单元格样式
     *
     * @param cellClass 用于存储当前单元格的全部class
     * @param cellStyle POI单元格样式类
     * @author Kirito丶城
     * @date 2022/8/25
     */
    private static void createCellStyle(List<String> cellClass, CellStyle cellStyle) {
        // 对齐方式
        switch (cellStyle.getAlignment()) {
            case CENTER:
                cellClass.add(TEXT_ALIGN_CENTER_CLASS);
                break;
            case RIGHT:
                cellClass.add(TEXT_ALIGN_RIGHT_CLASS);
                break;

            default:
                break;
        }
    }

    /**
     * 添加单元格字体样式
     *
     * @param cssMap    存储全部自定义样式内容。 key-class名  value-对应样式
     * @param font      字体对象
     * @param cellClass 用于存储当前单元格的全部class
     * @author Kirito丶城
     * @date 2022/8/25
     */
    private static void addFontStyle(Map<String, String> cssMap, Font font, List<String> cellClass) {
        // 字体大小，默认字体大小为16，如果字体超过16则生成字体大小样式
        double fontSize = (font.getFontHeightInPoints() / 72.0) * 96;
        if (fontSize >= 15) {
            String fontSizeStr = DECIMAL_FORMAT.format(fontSize);
            String className = FONT_SIZE_CLASS_PREFIX + fontSizeStr;
            cssMap.put(className, FONT_SIZE + fontSizeStr + PX_SEMICOLON);
            cellClass.add(className);
        }

        // 字体加粗
        if (font.getBold()) {
            cellClass.add(BOLD_CLASS);
        }

        // 字体倾斜
        if (font.getItalic()) {
            cellClass.add(ITALIC_CLASS);
        }

        // 删除线
        if (font.getStrikeout()) {
            cellClass.add(STRIKEOUT_CLASS);
        }

        // 下划线
        if (font.getUnderline() > 0) {
            cellClass.add(UNDERLINE_CLASS);
        }
    }

    /**
     * 生成列宽度代码
     *
     * @param tableCssStyle 表格样式字符串
     * @param st            sheet页
     * @param columnWidth   列宽存储map
     * @author Kirito丶城
     * @date 2022/8/25
     */
    private static void appendColWidthCss(StringBuilder tableCssStyle, int st, Map<Integer, Float> columnWidth) {
        // 生成列宽度代码
        Set<Integer> keySet = columnWidth.keySet();
        for (Integer key : keySet) {
            String colWidthCss = CLASS_NAME_SHEET +
                    (st + 1) +
                    " table tr td:nth-child(" +
                    key +
                    BRACKETS_RIGHT_AND_CURLY_BRACES_LEFT +
                    WIDTH_COLON +
                    DECIMAL_FORMAT.format(columnWidth.get(key)) +
                    PX_SEMICOLON +
                    CURLY_BRACES_RIGHT;
            tableCssStyle.append(colWidthCss);
        }
    }

    /**
     * 添加自定义样式
     *
     * @param style  表格样式字符串
     * @param cssMap 存储全部自定义样式内容。 key-class名  value-对应样式
     * @author Kirito丶城
     * @date 2022/8/25
     */
    private static void appendCss(StringBuilder style, Map<String, String> cssMap) {
        if (MapUtils.isEmpty(cssMap)) {
            return;
        }
        Set<String> classNames = cssMap.keySet();
        if (CollectionUtils.isEmpty(classNames)) {
            return;
        }
        for (String className : classNames) {
            String statement = cssMap.get(className);
            if (TextUtils.isEmpty(statement)) {
                continue;
            }
            style.append(CLASS_SELECTOR).append(className).append(CURLY_BRACES_LEFT).append(statement).append(CURLY_BRACES_RIGHT);
        }
    }

    /**
     * 修改表格数据，清除仅有样式没有值的多余列和行
     *
     * @param tableData   修剪前的表格数据
     * @param columnWidth 修剪前的列宽数据，执行后也是修剪后的数据
     * @return java.util.List<java.util.List < java.lang.Object>> 修剪后的表格数据
     * @author Kirito丶城
     * @date 2022/8/26
     */
    private static List<List<Object>> trimTableData(List<List<Object>> tableData, Map<Integer, Float> columnWidth) {
        if (CollectionUtils.isEmpty(tableData)) {
            return new ArrayList<>();
        }
        // 最大单元格长度，用于判断清除掉的列宽样式
        int maxCellSize = 0;
        // 不为空行最大行的index
        int maxRowNotNullIndex = 0;
        // 值不为null最大单元格的index
        int maxCellNotNullIndex = 0;
        int tableDataSize = tableData.size();
        // 循环查找出每行的最大列，与最大行
        for (int i = 0; i < tableData.size(); i++) {
            List<Object> rowData = tableData.get(i);
            if (CollectionUtils.isEmpty(rowData)) {
                continue;
            }
            // 当前行最后一个有值单元格的列
            int rowMaxCellIndex = -1;
            int rowDataSize = rowData.size();
            for (int j = 0; j < rowDataSize; j++) {
                Object val = rowData.get(j);
                if (val != null && !"".equals(val)) {
                    rowMaxCellIndex = j;
                }
            }
            maxCellSize = Math.max(maxCellSize, rowDataSize);
            maxCellNotNullIndex = Math.max(maxCellNotNullIndex, rowMaxCellIndex);
            // 如果当前行有值，则代表非空行，记录下标准备裁剪
            if (rowMaxCellIndex > -1) {
                maxRowNotNullIndex = i;
            }
        }
        // 如果存在空行则裁剪掉
        if (maxRowNotNullIndex + 1 < tableDataSize) {
            tableData = tableData.subList(0, maxRowNotNullIndex + 1);
        }
        // 删除多余的列
        List<List<Object>> list = new ArrayList<>();
        for (List<Object> rowData : tableData) {
            if (CollectionUtils.isEmpty(rowData)) {
                continue;
            }
            if (maxCellNotNullIndex + 1 < rowData.size()) {
                List<Object> subList = rowData.subList(0, maxCellNotNullIndex + 1);
                list.add(subList);
            } else {
                list.add(rowData);
            }
        }
        // 清除掉无用的列宽度样式
        for (int i = maxCellNotNullIndex + 2; i <= maxCellSize; i++) {
            columnWidth.remove(i);
        }
        return list;
    }

    /**
     * 生成单元格class代码段
     *
     * @param cellClass 单元格样式数据
     * @return java.lang.String
     * @author Kirito丶城
     * @date 2022/8/26
     */
    private static String createCellClass(List<String> cellClass) {
        StringBuilder classHtml = new StringBuilder();
        // 如果存在样式，则生成class
        if (CollectionUtils.isNotEmpty(cellClass)) {
            classHtml.append(" class='");
            for (String className : cellClass) {
                classHtml.append(className).append(" ");
            }
            classHtml.deleteCharAt(classHtml.length() - 1);
            classHtml.append("'");
        }
        return classHtml.toString();
    }

    /**
     * 创建默认的单元格HTML(没有合并的单元格)
     *
     * @param rowCellClass 当前行内全部的单元格
     * @param rowValues    当前行内全部的数据
     * @return java.lang.String
     * @author Kirito丶城
     * @date 2022/8/26
     */
    private static String createDefaultTdHtml(List<List<String>> rowCellClass, List<Object> rowValues) {
        StringBuilder tdHtml = new StringBuilder();
        int rowCellClassSize = 0;
        // 如果当行有数据则获取有集合长度
        if (CollectionUtils.isNotEmpty(rowCellClass)) {
            rowCellClassSize = rowCellClass.size();
        }
        for (int j = 0; j < rowValues.size(); j++) {
            Object rowValue = rowValues.get(j);
            tdHtml.append("<td");
            // 如果没有下标越界，则获取列样式
            if (rowCellClassSize > j && rowValue != null && !"".equals(rowValue)) {
                List<String> cellClass = rowCellClass.get(j);
                // 如果存在样式，则生成class
                tdHtml.append(createCellClass(cellClass));
            }
            tdHtml.append(">");
            tdHtml.append(rowValue);
            tdHtml.append("</td>");
        }
        return tdHtml.toString();
    }

    /**
     * 生成表格HTML
     *
     * @param tableData         表格数据
     * @param tableRowCellClass 表格内所有行的class
     * @param hasMerged         是否有合并单元格
     * @param mergedRegions     合并单元格的坐标
     * @return java.lang.String 整个表格HTML
     * @author Kirito丶城
     * @date 2022/8/26
     */
    private static String createTableHtml(List<List<Object>> tableData, List<List<List<String>>> tableRowCellClass, boolean hasMerged, List<String> mergedRegions) {
        int tableRowCellClassSize = tableRowCellClass.size();
        StringBuilder rowHtml = new StringBuilder("<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">");
        for (int i = 0; i < tableData.size(); i++) {
            // 当没有超过存储最大行，那么就获取当前行的全部样式
            List<List<String>> rowCellClass = null;
            if (tableRowCellClassSize > i) {
                rowCellClass = tableRowCellClass.get(i);
            }
            List<Object> rowValues = tableData.get(i);
            rowHtml.append("<tr>");
            if (hasMerged) {
                for (int j = 0; j < rowValues.size(); j++) {
                    Object rowValue = rowValues.get(j);
                    if (!INVALID_CELL.equals(rowValue)) {
                        // 去除已经格式化完成单元格
                        if (mergedRegions.contains(i + ":" + j)) {
                            rowHtml.append(rowValue);
                        } else {
                            int rowCellClassSize = 0;
                            // 如果当行有数据则获取有集合长度
                            if (CollectionUtils.isNotEmpty(rowCellClass)) {
                                rowCellClassSize = rowCellClass.size();
                            }
                            rowHtml.append("<td");
                            // 如果没有下标越界，则获取列样式
                            if (rowCellClassSize > j && rowValue != null && !"".equals(rowValue)) {
                                List<String> cellClass = rowCellClass.get(j);
                                // 如果存在样式，则生成class
                                rowHtml.append(createCellClass(cellClass));
                            }
                            rowHtml.append(">");
                            rowHtml.append(rowValue);
                            rowHtml.append("</td>");
                        }
                    }
                }
            } else {
                rowHtml.append(createDefaultTdHtml(rowCellClass, rowValues));
            }
            rowHtml.append("</tr>");
        }
        return rowHtml.append("</table>").toString();
    }

    /**
     * 获取表格合并坐标
     *
     * @param sheet             Sheet对象
     * @param tableData         表格数据
     * @param tableRowCellClass 表格内所有行的class
     * @return java.util.List<java.lang.String>
     * @author Kirito丶城
     * @date 2022/8/26
     */
    private static List<String> getMergedRegions(Sheet sheet, List<List<Object>> tableData, List<List<List<String>>> tableRowCellClass) {
        List<String> mergedRegions = new ArrayList<>();
        if (sheet.getNumMergedRegions() > 0) {
            List<CellRangeAddress> cellRangeAddresses = sheet.getMergedRegions();
            for (CellRangeAddress cellRangeAddress : cellRangeAddresses) {
                int firstRow = cellRangeAddress.getFirstRow();
                int lastRow = cellRangeAddress.getLastRow();
                int firstColumn = cellRangeAddress.getFirstColumn();
                int lastColumn = cellRangeAddress.getLastColumn();
                int rowspan = lastRow - firstRow + 1;
                int colspan = lastColumn - firstColumn + 1;
                String tdProperty = " rowspan='" + rowspan + "' colspan='" + colspan + "'";
                // 去除合并单元格后的多余的<td> 并将有合并行的HTML生成
                for (int i = firstRow; i <= lastRow; i++) {
                    // 去除合并单元格后的多余的<td>
                    List<Object> rowValues = tableData.get(i);
                    for (int j = firstColumn; j < lastColumn; j++) {
                        rowValues.set(j + 1, INVALID_CELL);
                    }
                    if (rowspan > 1) {
                        // 当合并行时第一行的<td>数量是不变的 需要加上rowspan属性，其他被合并的行内的<td>则会减少列合并数，如果列没有合并那么就会在对应位置减少1个<td>
                        if (i > firstRow) {
                            tableData.get(i).set(firstColumn, INVALID_CELL);
                        }
                    }
                }
                List<Object> rowValues = tableData.get(firstRow);
                StringBuffer tdHtml = new StringBuffer();
                // 寻找合并单元格第一个<td> 也就是需要些属性的<td>
                tdHtml.append("<td");
                if (CollectionUtils.isNotEmpty(tableRowCellClass) && tableRowCellClass.size() > firstRow) {
                    List<List<String>> rowCellClass = tableRowCellClass.get(firstRow);
                    if (CollectionUtils.isNotEmpty(rowCellClass) && rowCellClass.size() > firstColumn) {
                        List<String> cellClass = rowCellClass.get(firstColumn);
                        // 如果存在样式，则生成class
                        tdHtml.append(createCellClass(cellClass));
                    }
                }
                tdHtml.append(tdProperty);
                tdHtml.append(">");
                tdHtml.append(rowValues.get(firstColumn));
                tdHtml.append("</td>");
                rowValues.set(firstColumn, tdHtml);
                mergedRegions.add(firstRow + ":" + firstColumn);
            }
        }
        return mergedRegions;
    }

    /**
     * 当单元格的值写入，rowValues中
     *
     * @param rowValues 用于承接值的集合
     * @param cell      单元格对象
     * @author Kirito丶城
     * @date 2022/8/29
     */
    private static void setValues(List<Object> rowValues, Cell cell) {
        int columnIndex = cell.getColumnIndex();
        switch (cell.getCellType()) {
            case STRING:

                rowValues.set(columnIndex, cell.getStringCellValue());
                break;
            case BOOLEAN:

                rowValues.set(columnIndex, cell.getBooleanCellValue());
                break;
            case FORMULA:

                switch (cell.getCachedFormulaResultType()) {
                    case NUMERIC:

                        rowValues.set(columnIndex, cell.getNumericCellValue());
                        break;

                    case STRING:
                        rowValues.set(columnIndex, cell.getStringCellValue());
                        break;

                    case BOOLEAN:

                        rowValues.set(columnIndex, cell.getBooleanCellValue());
                        break;

                    default:
                        rowValues.set(columnIndex, "");
                        break;
                }
                break;
            case NUMERIC:

                double cellValue = cell.getNumericCellValue();
                DecimalFormat df = new DecimalFormat("#");
                String strIntegerVal = df.format(cellValue);
                double integer = Double.parseDouble(strIntegerVal);
                // 如果截取后的值与获取出的值相同，那么说明为整数
                rowValues.set(columnIndex, cellValue == integer ? strIntegerVal : cellValue);
                break;

            default:
                rowValues.set(columnIndex, "");
                break;
        }
    }

    /**
     * flush close FileWriter
     *
     * @param htmlBufferedWriter html文件
     * @param cssFileWriter      css文件
     * @author Kirito丶城
     * @date 2022/8/29
     */
    private static void closeFileWriter(BufferedWriter htmlBufferedWriter, FileWriter cssFileWriter) {
        // 只关闭 htmlBufferedWriter 即可，其参数流会自动关闭
        if (htmlBufferedWriter != null) {
            try {
                htmlBufferedWriter.flush();
            } catch (IOException e) {
            }
            try {
                htmlBufferedWriter.close();
            } catch (IOException e) {
            }
        }
        if (cssFileWriter != null) {
            try {
                cssFileWriter.flush();
            } catch (IOException e) {
            }
            try {
                cssFileWriter.close();
            } catch (IOException e) {
            }
        }
    }
}
