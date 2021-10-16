/*
 * Copyright (C) 2017 zhouzhuo810
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package me.zhouzhuo.zzexcelcreator;

import android.util.SparseIntArray;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @author zhouzhuo810
 * Created by zz on 2017/1/16.
 */
public class ZzExcelCreator implements ExcelManager {

    private static ZzExcelCreator creator;
    private WritableWorkbook writableWorkbook;
    private WritableSheet writableSheet;
    private SparseIntArray maxColWidthArray;
    private SparseIntArray maxRowHeightArray;

    private ZzExcelCreator() {
        maxColWidthArray = new SparseIntArray();
        maxRowHeightArray = new SparseIntArray();
    }

    public static ZzExcelCreator getInstance() {
        if (creator == null) {
            synchronized (ZzExcelCreator.class) {
                if (creator == null) {
                    creator = new ZzExcelCreator();
                }
            }
        }
        return creator;
    }


    @Override
    public ZzExcelCreator createExcel(String pathDir, String name) throws IOException {
        File dir = new File(pathDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        writableWorkbook = Workbook.createWorkbook(new File(pathDir + File.separator + name + ".xls"));
        return this;
    }

    @Override
    public ZzExcelCreator openExcel(File file) throws IOException, BiffException {
        FileInputStream fis = new FileInputStream(file);
        Workbook wb = Workbook.getWorkbook(fis);
        writableWorkbook = Workbook.createWorkbook(file, wb);
        return this;
    }

    @Override
    public ZzExcelCreator createSheet(String name) {
        checkNullFirst();
        writableSheet = writableWorkbook.createSheet(name, 0);
        return this;
    }

    @Override
    public ZzExcelCreator openSheet(int position) {
        checkNullFirst();
        writableSheet = writableWorkbook.getSheet(position);
        checkNullArray();
        maxColWidthArray.clear();
        maxRowHeightArray.clear();
        return this;
    }

    @Override
    public ZzExcelCreator insertColumn(int position) {
        checkNullFirst();
        checkNullSecond();
        writableSheet.insertColumn(position);
        return this;
    }

    @Override
    public ZzExcelCreator insertRow(int position) {
        checkNullFirst();
        checkNullSecond();
        writableSheet.insertRow(position);
        return this;
    }

    @Override
    public ZzExcelCreator merge(int col1, int row1, int col2, int row2) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        writableSheet.mergeCells(col1, row1, col2, row2);
        return this;
    }

    @Override
    public ZzExcelCreator mergeColumn(int row, int col1, int col2) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        writableSheet.mergeCells(col1, row, col2, row);
        return this;
    }

    @Override
    public ZzExcelCreator mergeRow(int col, int row1, int row2) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        writableSheet.mergeCells(col, row1, col, row2);
        return this;
    }

    @Override
    public ZzExcelCreator fillNumber(int col, int row, double number, WritableCellFormat format) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        if (format != null && format.getWrap()) {
            setRowHeight(row, getRealRowHeight(row, number + "", format));
            setColumnWidth(col, getRealColWidth(col, number + "", format));
        }
        if (format == null) {
            writableSheet.addCell(new Number(col, row, number));
        } else {
            writableSheet.addCell(new Number(col, row, number, format));
        }
        return this;
    }

    @Override
    public ZzExcelCreator fillContent(int col, int row, String content, WritableCellFormat format) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        if (content == null) {
            content = "";
        }

        if (format != null && format.getWrap()) {
            setRowHeight(row, getRealRowHeight(row, content, format));
            setColumnWidth(col, getRealColWidth(col, content, format));
        }

        if (format == null) {
            writableSheet.addCell(new Label(col, row, content));
        } else {
            writableSheet.addCell(new Label(col, row, content, format));
        }

        return this;
    }
    
    @Override
    public ZzExcelCreator fillImage(int col, int row, double width, double height, File imgFile) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        writableSheet.addImage(new WritableImage(col, row, width, height, imgFile));
        return this;
    }
    
    @Override
    public ZzExcelCreator fillImage(int col, int row, double width, double height, byte[] imgBytes) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        writableSheet.addImage(new WritableImage(col, row, width, height, imgBytes));
        return this;
    }
    
    /**
     * 根据同一列的其他单元格宽度，综合获取列宽
     *
     * @param col     行号
     * @param content 内容
     * @param format  WritableCellFormat
     * @return 行高
     */
    private int getRealColWidth(int col, String content, WritableCellFormat format) {
        int realContentWidth = getRealContentWidth(content, format);
        int limitMaxWidth = ZzFormatCreator.getInstance().getMaxWidth();
        if (maxColWidthArray.indexOfKey(col) >= 0) {
            int max = maxColWidthArray.get(col);
            if (realContentWidth > max) {
                if (realContentWidth > limitMaxWidth) {
                    maxColWidthArray.put(col, limitMaxWidth);
                    return limitMaxWidth;
                } else {
                    maxColWidthArray.put(col, realContentWidth);
                    return realContentWidth;
                }
            } else {
                return max;
            }
        } else {
            if (realContentWidth > limitMaxWidth) {
                maxColWidthArray.put(col, limitMaxWidth);
                return limitMaxWidth;
            } else {
                maxColWidthArray.put(col, realContentWidth);
                return realContentWidth;
            }
        }
    }

    /**
     * 获取单元格内容宽度（对比每行宽度，取最大值）
     *
     * @param content 内容
     * @param format  WritableCellFormat
     * @return 宽度
     */
    private int getRealContentWidth(String content, WritableCellFormat format) {
        if (content != null) {
            int fontSize = format.getFont().getPointSize();
            float scale = fontSize * 1.0f / 14;
            if (content.contains("\n")) {
                String[] split = content.split("\n");
                int maxWidth = 0;
                for (String s : split) {
                    int chineseLength = getChineseNum(s);
                    int curWidth = (int) ((int) ((s.length() - chineseLength) * 1.15 + 2 + chineseLength * 3 + 0.5) * scale);
                    if (maxWidth < curWidth) {
                        maxWidth = curWidth;
                    }
                }
                return maxWidth;
            } else {
                int chineseLength = getChineseNum(content);
                return (int) ((int) ((content.length() - chineseLength) * 1.15 + 2 + chineseLength * 3 + 0.5) * scale);
            }
        }
        return 0;
    }


    /**
     * 根据同一行的其他单元格宽度，综合获取行高
     *
     * @param row     行号
     * @param content 内容
     * @return 行高
     */
    private int getRealRowHeight(int row, String content, WritableCellFormat format) {
        int realContentHeight = getRealContentHeight(content, format);
        if (maxRowHeightArray.indexOfKey(row) >= 0) {
            int max = maxRowHeightArray.get(row);
            if (realContentHeight > max) {
                maxRowHeightArray.put(row, realContentHeight);
                return realContentHeight;
            } else {
                return max;
            }
        } else {
            maxRowHeightArray.put(row, realContentHeight);
            return realContentHeight;
        }
    }

    /**
     * 获取单元格内容高度（所占行数*单行高度）
     *
     * @param content 内容
     * @return 高度
     */
    private int getRealContentHeight(String content, WritableCellFormat format) {
        if (content != null) {
            int lineCount = 0;
            if (content.contains("\n")) {
                String[] split = content.split("\n");
                lineCount += split.length;
                for (String s : split) {
                    int chineseLength = getChineseNum(s);
                    int curWidth = (int) ((s.length() - chineseLength) * 1.15 + 2 + chineseLength * 3 + 0.5);
                    if (curWidth > ZzFormatCreator.getInstance().getMaxWidth()) {
                        lineCount += (curWidth / ZzFormatCreator.getInstance().getMaxWidth());
                    }
                }
            } else {
                int chineseLength = getChineseNum(content);
                int curWidth = (int) ((content.length() - chineseLength) * 1.15 + 2 + chineseLength * 3 + 0.5);
                if (curWidth > ZzFormatCreator.getInstance().getMaxWidth()) {
                    lineCount += (curWidth / ZzFormatCreator.getInstance().getMaxWidth() + 1);
                } else {
                    lineCount++;
                }
            }
            return lineCount * format.getFont().getPointSize() * 30;

        }
        return format.getFont().getPointSize() * 30;
    }

    /**
     * 统计content中是汉字的个数
     *
     * @param content 内容
     * @return 个数
     */
    private int getChineseNum(String content) {
        int lenOfChinese = 0;
        Pattern p = Pattern.compile("[\u4e00-\u9fa5]");    //汉字的Unicode编码范围
        Matcher m = p.matcher(content);
        while (m.find()) {
            lenOfChinese++;
        }
        return lenOfChinese;
    }

    @Override
    public ZzExcelCreator setRowHeight(int position, int height) throws RowsExceededException {
        checkNullFirst();
        checkNullSecond();
        writableSheet.setRowView(position, height);
        return this;
    }

    @Override
    public ZzExcelCreator setRowHeightFromTo(int from, int to, int height) throws RowsExceededException {
        checkNullFirst();
        checkNullSecond();
        for (int i = from; i < to; i++) {
            writableSheet.setRowView(i, height);
        }
        return this;
    }
    
    @Override
    public ZzExcelCreator setRowHeightLikeWPS(int position, int height) throws RowsExceededException {
        checkNullFirst();
        checkNullSecond();
        writableSheet.setRowView(position, height * 20);
        return this;
    }
    
    @Override
    public ZzExcelCreator setRowHeightLikeWPSFromTo(int from, int to, int height) throws RowsExceededException {
        checkNullFirst();
        checkNullSecond();
        for (int i = from; i < to; i++) {
            writableSheet.setRowView(i, height * 20);
        }
        return this;
    }
    
    @Override
    public ZzExcelCreator setColumnWidth(int position, int width) {
        checkNullFirst();
        checkNullSecond();
        writableSheet.setColumnView(position, width);
        return this;
    }

    @Override
    public ZzExcelCreator setColumnWidth(int from, int to, int width) {
        checkNullFirst();
        checkNullSecond();
        for (int i = from; i < to; i++) {
            writableSheet.setColumnView(i, width);
        }
        return this;
    }

    @Override
    public ZzExcelCreator deleteColumn(int position) {
        checkNullFirst();
        checkNullSecond();
        writableSheet.removeColumn(position);
        return this;
    }

    @Override
    public ZzExcelCreator deleteRow(int position) {
        checkNullFirst();
        checkNullSecond();
        writableSheet.removeRow(position);
        return this;
    }

    @Override
    public String getCellContent(int col, int row) {
        checkNullFirst();
        checkNullSecond();
        return writableSheet.getWritableCell(col, row).getContents();
    }

    @Override
    public void close() throws IOException, WriteException {
        if (writableWorkbook != null) {
            writableWorkbook.write();
            writableWorkbook.close();
        }
        writableWorkbook = null;
        writableSheet = null;
    }

    private void checkNullFirst() {
        if (writableWorkbook == null) {
            throw new NullPointerException("writableWorkbook is null, please invoke the #createExcel(String, String) method or the #openExcel(File) method first.");
        }
    }

    private void checkNullSecond() {
        if (writableSheet == null) {
            throw new NullPointerException("writableSheet is null, please invoke the #createSheet(String) method or the #openSheet(int) method first.");
        }
    }

    private void checkNullArray() {
        if (maxRowHeightArray == null) {
            maxRowHeightArray = new SparseIntArray();
        }
        if (maxColWidthArray == null) {
            maxColWidthArray = new SparseIntArray();
        }
    }

    protected WritableSheet getWritableSheet() {
        return writableSheet;
    }

    protected WritableWorkbook getWritableWorkbook() {
        return writableWorkbook;
    }
}
