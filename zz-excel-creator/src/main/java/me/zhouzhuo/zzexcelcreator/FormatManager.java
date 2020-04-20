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
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

/**
 * @author zhouzhuo810
 * Created by zz on 2017/1/16.
 */
public interface FormatManager {
    
    /**
     * 创建字体
     *
     * @param fontName 字体名字 #WritableFont.ARIAL
     * @return ZzFormatCreator
     * @throws WriteException ex
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
     * 设置粗体
     *
     * @param bold 是否加粗
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setFontBold(boolean bold) throws WriteException;
    
    /**
     * 设置下划线
     *
     * @param underline 是否下划线
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setUnderline(boolean underline) throws WriteException;
    
    /**
     * 设置双重下划线
     *
     * @param doubleUnderline 是否双重下划线
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setDoubleUnderline(boolean doubleUnderline) throws WriteException;
    
    /**
     * 设置是否斜体
     *
     * @param italic 是否斜体
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setItalic(boolean italic) throws WriteException;
    
    /**
     * 设置是否自适应宽高
     *
     * @param wrap     是否自适应
     * @param maxWidth 最大宽度，到达该宽度则换行。
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setWrapContent(boolean wrap, int maxWidth) throws WriteException;
    
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
     * 设置单元格边框样式，可调用多次，如果一个单元格需要上和下边框，调用两次即可
     *
     * @param border          边框位置
     * @param borderLineStyle 边框线样式
     * @param colour          边框线颜色
     * @return ZzFormatCreator
     * @throws WriteException ex
     */
    ZzFormatCreator setBorder(Border border, BorderLineStyle borderLineStyle, Colour colour) throws WriteException;
    
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
    
    /**
     * 获取最大宽度
     *
     * @return 最大宽度
     */
    int getMaxWidth();
}
