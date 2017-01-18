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
public class ZzFormatCreator implements FormatManager {

    private WritableFont font;
    private WritableCellFormat cellFormat;
    private static ZzFormatCreator creator;

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
        return this;
    }

    @Override
    public ZzFormatCreator setFontColor(Colour color) throws WriteException {
        checkNull();
        font.setColour(color);
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

}
