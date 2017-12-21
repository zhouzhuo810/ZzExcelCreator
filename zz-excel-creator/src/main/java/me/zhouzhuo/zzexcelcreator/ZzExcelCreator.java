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

import android.util.Log;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @author zhouzhuo810
 *         Created by zz on 2017/1/16.
 */
public class ZzExcelCreator implements ExcelManager {

    private static ZzExcelCreator creator;
    private WritableWorkbook writableWorkbook;
    private WritableSheet writableSheet;

    private ZzExcelCreator() {
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
        if (!dir.exists())
            dir.mkdirs();
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
        Log.e("xxxx", writableWorkbook.getNumberOfSheets() + "");
        writableSheet = writableWorkbook.getSheet(position);
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
        if (format == null)
            writableSheet.addCell(new Number(col, row, number));
        else
            writableSheet.addCell(new Number(col, row, number, format));
        return this;
    }

    @Override
    public ZzExcelCreator fillContent(int col, int row, String content, WritableCellFormat format) throws WriteException {
        checkNullFirst();
        checkNullSecond();
        if (format == null)
            writableSheet.addCell(new Label(col, row, content));
        else
            writableSheet.addCell(new Label(col, row, content, format));
        return this;
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
        checkNullFirst();
        writableWorkbook.write();
        writableWorkbook.close();
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


}
