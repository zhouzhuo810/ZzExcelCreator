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

import java.io.File;
import java.io.IOException;

import jxl.read.biff.BiffException;
import jxl.write.WritableCellFormat;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @author zhouzhuo810
 *         Created by zz on 2017/1/16.
 */
public interface ExcelManager {

    /**
     * 创建Excel文件
     *
     * @param pathDir 文件夹地址
     * @param name    名字(不用带后缀,如果之后还要操作此文件，请使用英文或数字命名)
     * @return ZzExcelCreator
     * @throws IOException ex
     */
    ZzExcelCreator createExcel(String pathDir, String name) throws IOException;


    /**
     * 打开Excel文件
     *
     * @param file File对象
     * @return ZzExcelCreator
     * @throws IOException   ex
     * @throws BiffException ex
     */
    ZzExcelCreator openExcel(File file) throws IOException, BiffException;

    /**
     * 创建工作表
     *
     * @param name 工作表
     * @return ZzExcelCreator
     */
    ZzExcelCreator createSheet(String name);

    /**
     * 打开工作表
     *
     * @param position 工作表位置
     * @return ZzExcelCreator
     */
    ZzExcelCreator openSheet(int position);

    /**
     * 插入一个空列。
     *
     * @param position 目标位置
     * @return ZzExcelCreator
     */
    ZzExcelCreator insertColumn(int position);

    /**
     * 插入一个空行。
     *
     * @param position 目标位置
     * @return ZzExcelCreator
     */
    ZzExcelCreator insertRow(int position);

    /**
     * 合并区域
     *
     * @param col1 区域左上角的列号
     * @param row1 区域左上角的行号
     * @param col2 区域右下角的列表
     * @param row2 区域右下角的行号
     * @return ZzExcelCreator
     * @throws WriteException ex
     */
    ZzExcelCreator merge(int col1, int row1, int col2, int row2) throws WriteException;


    /**
     * 合并列(水平合并)
     *
     * @param row  行号
     * @param col1 起始列号
     * @param col2 结束列号
     * @return ZzExcelCreator
     * @throws WriteException ex
     */
    ZzExcelCreator mergeColumn(int row, int col1, int col2) throws WriteException;


    /**
     * 合并行(垂直合并)
     *
     * @param col  列号
     * @param row1 起始行号
     * @param row2 结束行号
     * @return ZzExcelCreator
     * @throws WriteException ex
     */
    ZzExcelCreator mergeRow(int col, int row1, int row2) throws WriteException;


    /**
     * 填入表格内容(数字)
     *
     * @param col    列号
     * @param row    行号
     * @param number 数字
     * @param format 格式(默认时传null即可)
     * @return ZzExcelCreator
     * @throws WriteException ex
     */
    ZzExcelCreator fillNumber(int col, int row, double number, WritableCellFormat format) throws WriteException;

    /**
     * 填充表格内容(字符串)
     *
     * @param col     列号
     * @param row     行号
     * @param content 要填充的内容
     * @param format  格式(默认时传null即可)
     * @return ZzExcelCreator
     * @throws WriteException ex
     */
    ZzExcelCreator fillContent(int col, int row, String content, WritableCellFormat format) throws WriteException;


    /**
     * 设置行高
     * <p>
     * <p>(推荐行高为400-500)</p>
     *
     * @param position 行号
     * @param height   行高
     * @return ZzExcelCreator
     * @throws RowsExceededException ex
     */
    ZzExcelCreator setRowHeight(int position, int height) throws RowsExceededException;

    /**
     * 设置行高
     * <p>
     * <p>(推荐行高为400-500)</p>
     *
     * @param from   起始行号
     * @param to     目标行号+1
     * @param height 行高
     * @return ZzExcelCreator
     * @throws RowsExceededException ex
     */
    ZzExcelCreator setRowHeightFromTo(int from, int to, int height) throws RowsExceededException;

    /**
     * 设置列宽
     * <p>
     * <p>(推荐列宽为20-30)</p>
     *
     * @param position 列号
     * @param width    列宽
     * @return ZzExcelCreator ex
     */
    ZzExcelCreator setColumnWidth(int position, int width);


    /**
     * 设置列宽
     * <p>
     * <p>(推荐列宽为20-30)</p>
     *
     * @param from  起始列号
     * @param to    目标列号+1
     * @param width 列宽
     * @return ZzExcelCreator
     */
    ZzExcelCreator setColumnWidth(int from, int to, int width);

    /**
     * 删除某列表
     *
     * @param position 目标位置
     * @return ZzExcelCreator
     */
    ZzExcelCreator deleteColumn(int position);

    /**
     * 删除某行
     *
     * @param position 目标位置
     * @return ZzExcelCreator
     */
    ZzExcelCreator deleteRow(int position);

    /**
     * 获取单元格内容
     *
     * @param col 列
     * @param row 行
     * @return 值
     */
    String getCellContent(int col, int row);

    /**
     * 结束操作
     *
     * @throws IOException    ex
     * @throws WriteException ex
     */
    void close() throws IOException, WriteException;

}
