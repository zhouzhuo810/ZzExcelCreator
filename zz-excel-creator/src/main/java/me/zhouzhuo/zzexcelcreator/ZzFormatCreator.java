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

import java.security.InvalidParameterException;

import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

/**
 * @author zhouzhuo810
 * Created by zz on 2017/1/16.
 */
public class ZzFormatCreator implements FormatManager {
    
    private WritableFont font;
    private WritableCellFormat cellFormat;
    private static ZzFormatCreator creator;
    
    private int maxWidth = 100;
    
    private ZzFormatCreator() {
    }
    
    public static ZzFormatCreator getInstance() {
        if (creator == null) {
            synchronized (ZzFormatCreator.class) {
                if (creator == null) {
                    creator = new ZzFormatCreator();
                }
            }
        }
        return creator;
    }
    
    public static ZzFormatCreator newInstance() {
        return new ZzFormatCreator();
    }
    
    @Override
    public ZzFormatCreator createCellFont(WritableFont.FontName fontName) throws WriteException {
        font = new WritableFont(fontName);
        font.setPointSize(12);
        font.setColour(Colour.BLACK);
        cellFormat = new WritableCellFormat();
        cellFormat.setFont(font);
        cellFormat.setAlignment(Alignment.CENTRE);
        cellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
        cellFormat.setWrap(true);
        return this;
    }
    
    @Override
    public ZzFormatCreator setFontColor(Colour color) throws WriteException {
        checkNull();
        font.setColour(color);
        return this;
    }
    
    @Override
    public ZzFormatCreator setFontBold(boolean bold) throws WriteException {
        checkNull();
        font.setBoldStyle(bold ? WritableFont.BOLD : WritableFont.NO_BOLD);
        return this;
    }
    
    @Override
    public ZzFormatCreator setUnderline(boolean underline) throws WriteException {
        checkNull();
        font.setUnderlineStyle(underline ? UnderlineStyle.SINGLE : UnderlineStyle.NO_UNDERLINE);
        return this;
    }
    
    @Override
    public ZzFormatCreator setDoubleUnderline(boolean doubleUnderline) throws WriteException {
        checkNull();
        font.setUnderlineStyle(doubleUnderline ? UnderlineStyle.DOUBLE : UnderlineStyle.NO_UNDERLINE);
        return this;
    }
    
    @Override
    public ZzFormatCreator setItalic(boolean italic) throws WriteException {
        checkNull();
        font.setItalic(italic);
        return this;
    }
    
    @Override
    public ZzFormatCreator setWrapContent(boolean wrap, int maxWidth) throws WriteException {
        checkNull();
        if (maxWidth <= 0) {
            throw new InvalidParameterException("maxWidth must > 0");
        }
        cellFormat.setWrap(wrap);
        this.maxWidth = maxWidth;
        return this;
    }
    
    @Override
    public ZzFormatCreator setAlignment(Alignment align, VerticalAlignment verticalAlign) throws WriteException {
        checkNull();
        cellFormat.setVerticalAlignment(verticalAlign);
        cellFormat.setAlignment(align);
        return this;
    }
    
    @Override
    public ZzFormatCreator setBackgroundColor(Colour color) throws WriteException {
        checkNull();
        cellFormat.setBackground(color);
        return this;
    }
    
    @Override
    public ZzFormatCreator setBorder(Border border, BorderLineStyle borderLineStyle, Colour colour) throws WriteException {
        checkNull();
        cellFormat.setBorder(border, borderLineStyle, colour);
        return this;
    }
    
    @Override
    public ZzFormatCreator setFontSize(int fontSize) throws WriteException {
        checkNull();
        font.setPointSize(fontSize);
        return this;
    }
    
    @Override
    public WritableCellFormat getCellFormat() {
        return cellFormat;
    }
    
    
    private void checkNull() {
        if (font == null) {
            throw new NullPointerException("WritableFont is null, Please invoke the method #createCellFont().");
        }
        if (cellFormat == null) {
            throw new NullPointerException("WritableCellFormat is null, Please invoke the method #createCellFont().");
        }
    }
    
    @Override
    public int getMaxWidth() {
        return maxWidth;
    }
}
