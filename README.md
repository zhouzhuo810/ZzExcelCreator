# ZzExcelCreator
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
compile 'me.zhouzhuo.zzexcelcreator:zz-excel-creator:1.0.0'
```


### 创建Excel文件和工作表

```java
                    ZzExcelCreator
                            .getInstance()
                            .createExcel(PATH, params[0])  //生成excel文件
                            .createSheet(params[1])        //生成sheet工作表
                            .close();
```
### 打开Excel文件和工作表
```java
                            ZzExcelCreator
                                    .getInstance()
                                    .openExcel(new File(PATH + fileName + ".xls"))  //打开Excel文件
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
                                    .getCellFormat();
```

### 设置行高、列宽和写入字符串或数字

```java
                            ZzExcelCreator
                                    .getInstance()
                                    .openExcel(new File(PATH + fileName + ".xls"))  
                                    .openSheet(0)
                                    .setColumnWidth(Integer.parseInt(col), 25)   //设置列宽
                                    .setRowHeight(Integer.parseInt(row), 400)    //设置行高
                                    .fillContent(Integer.parseInt(col), Integer.parseInt(row), str, format)  //填入字符串
                                    .fillNumber(Integer.parseInt(col), Integer.parseInt(row), Double.parseDouble(str), format)  //填入数字
                                    .close();
```

### 最后就是，这些操作最好在子线程操作。
