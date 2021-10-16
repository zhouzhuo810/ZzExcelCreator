# ZzExcelCreator

[![](https://jitpack.io/v/zhouzhuo810/ZzExcelCreator.svg)](https://jitpack.io/#zhouzhuo810/ZzExcelCreator)

Excel表格生成工具

项目地址:https://github.com/zhouzhuo810/ZzExcelCreator
（欢迎star!）

效果图：

![excel1.jpg](http://upload-images.jianshu.io/upload_images/2788864-5c0055e4dcc74d5a.jpg?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
![excel2.jpg](http://upload-images.jianshu.io/upload_images/2788864-edff815539021b6a.jpg?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
![excel3.jpg](http://upload-images.jianshu.io/upload_images/2788864-3da6e6ab4db88b22.jpg?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

最近做项目用到jxl.jar来生成Excel表格；

但是发现jxl源码都没有注释的，方法也没有说明，
虽然最后在网上找到了对应的方法。

不过这不是我的style，果断自己封装一下，添加注释。


下面介绍一下用法：

### Gradle：

```
	allprojects {
		repositories {
			...
			maven { url 'https://jitpack.io' }
		}
	}
```

```
	dependencies {
	        implementation 'com.github.zhouzhuo810:ZzExcelCreator:1.0.9'
	}
```



### 创建Excel文件和工作表

```java
        ZzExcelCreator
                .getInstance()
                .createExcel(filePath, fileName)  //生成excel文件
                .createSheet(sheetName)        //生成sheet工作表
                .close();
```

### 添加工作表

```java
        ZzExcelCreator
                .getInstance()
                .openExcel(new File(filePath + fileName + ".xls"))  //如果不想覆盖文件，注意是openExcel
                .createSheet(sheetName)
                .close();
```

### 打开Excel文件和工作表
```java
        ZzExcelCreator
                .getInstance()
                .openExcel(new File(filePath + fileName + ".xls"))  //打开Excel文件
                .openSheet(0)                                   //打开Sheet工作表
                ... ...
                .close();
```

### 设置单元格内容格式：

```java
        //设置单元格内容格式
        WritableCellFormat format = ZzFormatCreator
                .getInstance()
                .createCellFont(WritableFont.ARIAL)  //设置字体
                .setAlignment(Alignment.CENTRE, VerticalAlignment.CENTRE)  //设置对齐方式(水平和垂直)
                .setFontSize(14)                    //设置字体大小
                .setFontColor(Colour.ROSE)          //设置字体颜色
                .setFontBold(true)                  //设置是否加粗，默认false
                .setUnderline(true)                 //设置是否画下划线，默认false
                //.setDoubleUnderline(true)         //设置是否画双重下划线，默认false,和setUnderline只有一个生效
                .setItalic(true)                    //设置是否斜体
                .setWrapContent(true, 100)          //设置是否自适应宽高，如果自适应，必须设置最大列宽（不能太大，否则可能无效）。
                .setBackgroundColor(ColourUtil.getCustomColor1("#99cc00"))  //设置单元格背景颜色，如果不设置边框，边框色会和背景色一致。
                .setBorder(Border.LEFT, BorderLineStyle.THIN, ColourUtil.getCustomColor2("#dddddd"))  //设置左边边框样式
                .setBorder(Border.TOP, BorderLineStyle.THIN, ColourUtil.getCustomColor2("#dddddd"))  //设置顶部边框样式
                .getCellFormat();
```

### 设置行高、列宽和写入字符串或数字

#### 写入数字

```java
        ZzExcelCreator
                .getInstance()
                .openExcel(new File(PATH + fileName + ".xls"))
                .openSheet(0)
                .setColumnWidth(colInt, 25)   //设置列宽(如果自适应宽度，代表内容字节的长度,即str.getBytes().length)
                .setRowHeight(rowInt, 400)    //设置行高，单位为磅的20倍
                .fillContent(colInt, rowInt, str, format)  //填入字符串
                .fillNumber(colInt, rowInt, Double.parseDouble(str), format)  //填入数字
                .close();
```

#### 写入字符串

```java
        ZzExcelCreator
                .getInstance()
                .openExcel(new File(PATH + fileName + ".xls"))
                .openSheet(0)
                .setColumnWidth(colInt, 25)      //设置列宽(如果自适应宽度，代表内容字节的长度,即str.getBytes().length)
                .setRowHeightLikeWPS(rowInt, 26) //设置行高同WPS，单位磅
                .fillContent(colInt, rowInt, str, format)  //填入字符串
                .fillNumber(colInt, rowInt, Double.parseDouble(str), format)  //填入数字
                .close();
```

#### 写入图片

```java
        ZzExcelCreator
                .getInstance()
                .openExcel(new File(PATH + fileName + ".xls"))
                .openSheet(0)
                .fillImage(Integer.parseInt(col), Integer.parseInt(row),
                        Double.parseDouble(width), Double.parseDouble(height),
                        new File(filePath))
                .close();
         //注意插入图片的宽高的单位为行高或列宽的倍数。
```


### 读取单元格内容

```java
        ZzExcelCreator zzExcelCreator = ZzExcelCreator
                .getInstance()
                .openExcel(new File(filePath + fileName + ".xls"))
                .openSheet(0);
        //读取单元格内容
        String content =  zzExcelCreator.getCellContent(colInt, rowInt);
        //别忘了close
        zzExcelCreator.close();
```

### 自定义颜色

```java
    //注意调用此方法必须保证ZzExcelCreator.getInstance().getWritableWorkbook()不为空。
    Colour colour1 = ColourUtil.getCustomColor1("#99cc00");
    Colour colour2 = ColourUtil.getCustomColor2("#99cc00");
    Colour colour3 = ColourUtil.getCustomColor3("#99cc00");
    Colour colour4 = ColourUtil.getCustomColor4("#99cc00");
    Colour colour5 = ColourUtil.getCustomColor5("#99cc00");
    Colour colour6 = ColourUtil.getCustomColor6("#99cc00");
```

### 注意一定别忘了close。

### 最后就是，这些操作最好在子线程操作。


