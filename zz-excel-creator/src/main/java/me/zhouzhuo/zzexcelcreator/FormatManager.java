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

import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

/**
 * @author zhouzhuo810
 *         Created by zz on 2017/1/16.
 */
public interface FormatManager {

    /**
     * 创建字体
     *
     * @param fontName 字体名字 #WritableFont.ARIAL
     * @return ZzFormatCreator
     * @throws WriteException  ex
     */
    ZzFormatCreator createCellFont(WritableFont.FontName fontName) throws WriteException;

    /**
     * 设置字体颜色
     *
     * @param color 颜色
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setFontColor(Colour color) throws WriteException;

    /**
     * 设置内容靠边或居中
     *
     * @param align         水平对其方式
     * @param verticalAlign 垂直对其方式
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setAlignment(Alignment align, VerticalAlignment verticalAlign) throws WriteException;

    /**
     * 设置背景颜色
     *
     * @param color 颜色
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setBackgroundColor(Colour color) throws WriteException;

    /**
     * 设置字体大小
     *
     * @param size 字体大小
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setFontSize(int size) throws WriteException;

    /**
     * 获取格式
     *
     * @return WritableCellFormat
     */
    WritableCellFormat getCellFormat();
}
