package me.zhouzhuo.zzexcelcreator;

import android.graphics.Color;

import jxl.format.Colour;

/**
 * jxl自定义颜色工具
 * Created by zz on 2018/11/6.
 */
public class ColourUtil {
    /**
     * 将十六进制颜色转换为jxl可用的颜色
     */
    public static Colour getCustomColour(String strColor) {
        int cl = Color.parseColor(strColor);
        Colour color = null;
        Colour[] colors = Colour.getAllColours();
        if ((colors != null) && (colors.length > 0)) {
            Colour crtColor = null;
            int[] rgb = null;
            int diff = 0;
            int minDiff = 999;
            for (int i = 0; i < colors.length; i++) {
                crtColor = colors[i];
                rgb = new int[3];
                rgb[0] = crtColor.getDefaultRGB().getRed();
                rgb[1] = crtColor.getDefaultRGB().getGreen();
                rgb[2] = crtColor.getDefaultRGB().getBlue();
                
                diff = Math.abs(rgb[0] - Color.red(cl))
                    + Math.abs(rgb[1] - Color.green(cl))
                    + Math.abs(rgb[2] - Color.blue(cl));
                if (diff < minDiff) {
                    minDiff = diff;
                    color = crtColor;
                }
            }
        }
        if (color == null)
            color = Colour.BLACK;
        return color;
    }
    
/*
    //上面的方法适用于Android平台， 如果是javaSE使用此方法。
    //import java.awt.*;
    public static Colour getCustomColour(String strColor) {
        Color cl = Color.decode(strColor);
        Colour color = null;
        Colour[] colors = Colour.getAllColours();
        if ((colors != null) && (colors.length > 0)) {
            Colour crtColor = null;
            int[] rgb = null;
            int diff = 0;
            int minDiff = 999;
            for (int i = 0; i < colors.length; i++) {
                crtColor = colors[i];
                rgb = new int[3];
                rgb[0] = crtColor.getDefaultRGB().getRed();
                rgb[1] = crtColor.getDefaultRGB().getGreen();
                rgb[2] = crtColor.getDefaultRGB().getBlue();
                
                diff = Math.abs(rgb[0] - cl.getRed())
                    + Math.abs(rgb[1] - cl.getGreen())
                    + Math.abs(rgb[2] - cl.getBlue());
                if (diff < minDiff) {
                    minDiff = diff;
                    color = crtColor;
                }
            }
        }
        if (color == null)
            color = Colour.BLACK;
        return color;
    }
    */
}
