package me.zhouzhuo.zzexcelcreator;

import android.graphics.Color;

import jxl.format.Colour;
import jxl.write.WritableWorkbook;

/**
 * jxl自定义颜色工具
 * 将一些用的比较少的Colour常量赋值为自定义的颜色，以达到自定义颜色效果。
 * Created by zz on 2018/11/6.
 */
public class ColourUtil {

    /**
     * 修改常量Colour.ROSE值为指定颜色
     *
     * @param colorStr 自定义颜色
     * @return Colour.ROSE
     */
    public static Colour getCustomColor1(String colorStr) {
        int color = Color.parseColor(colorStr); // 自定义的颜色
        WritableWorkbook workbook = ZzExcelCreator.getInstance().getWritableWorkbook();
        if (workbook == null) {
            throw new NullPointerException("Please invoke ZzExcelCreator.getInstance().createExcel() method first.");
        }
        workbook.setColourRGB(Colour.ROSE, Color.red(color), Color.green(color), Color.blue(color));
        return Colour.ROSE;
    }

    /**
     * 修改常量Colour.CORAL值为指定颜色
     *
     * @param colorStr 自定义颜色
     * @return Colour.CORAL
     */
    public static Colour getCustomColor2(String colorStr) {
        int color = Color.parseColor(colorStr); // 自定义的颜色
        WritableWorkbook workbook = ZzExcelCreator.getInstance().getWritableWorkbook();
        if (workbook == null) {
            throw new NullPointerException("Please invoke ZzExcelCreator.getInstance().createExcel() method first.");
        }
        workbook.setColourRGB(Colour.CORAL, Color.red(color), Color.green(color), Color.blue(color));
        return Colour.CORAL;
    }

    /**
     * 修改常量Colour.YELLOW值为指定颜色
     *
     * @param colorStr 自定义颜色
     * @return Colour.YELLOW
     */
    public static Colour getCustomColor3(String colorStr) {
        int color = Color.parseColor(colorStr); // 自定义的颜色
        WritableWorkbook workbook = ZzExcelCreator.getInstance().getWritableWorkbook();
        if (workbook == null) {
            throw new NullPointerException("Please invoke ZzExcelCreator.getInstance().createExcel() method first.");
        }
        workbook.setColourRGB(Colour.YELLOW, Color.red(color), Color.green(color), Color.blue(color));
        return Colour.YELLOW;
    }

    /**
     * 修改常量Colour.YELLOW2值为指定颜色
     *
     * @param colorStr 自定义颜色
     * @return Colour.YELLOW2
     */
    public static Colour getCustomColor4(String colorStr) {
        int color = Color.parseColor(colorStr); // 自定义的颜色
        WritableWorkbook workbook = ZzExcelCreator.getInstance().getWritableWorkbook();
        if (workbook == null) {
            throw new NullPointerException("Please invoke ZzExcelCreator.getInstance().createExcel() method first.");
        }
        workbook.setColourRGB(Colour.YELLOW2, Color.red(color), Color.green(color), Color.blue(color));
        return Colour.YELLOW2;
    }

    /**
     * 修改常量Colour.DARK_YELLOW值为指定颜色
     *
     * @param colorStr 自定义颜色
     * @return Colour.DARK_YELLOW
     */
    public static Colour getCustomColor5(String colorStr) {
        int color = Color.parseColor(colorStr); // 自定义的颜色
        WritableWorkbook workbook = ZzExcelCreator.getInstance().getWritableWorkbook();
        if (workbook == null) {
            throw new NullPointerException("Please invoke ZzExcelCreator.getInstance().createExcel() method first.");
        }
        workbook.setColourRGB(Colour.DARK_YELLOW, Color.red(color), Color.green(color), Color.blue(color));
        return Colour.DARK_YELLOW;
    }

    /**
     * 修改常量Colour.VERY_LIGHT_YELLOW值为指定颜色
     *
     * @param colorStr 自定义颜色
     * @return Colour.VERY_LIGHT_YELLOW
     */
    public static Colour getCustomColor6(String colorStr) {
        int color = Color.parseColor(colorStr); // 自定义的颜色
        WritableWorkbook workbook = ZzExcelCreator.getInstance().getWritableWorkbook();
        if (workbook == null) {
            throw new NullPointerException("Please invoke ZzExcelCreator.getInstance().createExcel() method first.");
        }
        workbook.setColourRGB(Colour.VERY_LIGHT_YELLOW, Color.red(color), Color.green(color), Color.blue(color));
        return Colour.VERY_LIGHT_YELLOW;
    }

}
