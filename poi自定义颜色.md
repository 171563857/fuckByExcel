HSSF针对的是2003版的excel文件
============================
XSSF针对的是2007版的excel文件
============================
颜色上限:
HSSF的调色板只有56种颜色
XSSF无上限

首先创建个HSSFWorkbook对象↓
```Java
HSSFWorkbook wb = new HSSFWorkbook();
```

我的做法是创建一个map对象用来存储样式↓
```Java
Map<String, HSSFCellStyle> color = new HashMap<>();
```

获取调色板↓
```Java
HSSFPalette palette = wb.getCustomPalette();
```

HSSF只能通过调色板自定义颜色,而调色板只能记录`56`种颜色,下标从8到64
可以查看类'PaletteRecord'能看到
PaletteRecord.FIRST_COLOR_INDEX  常量的值为8,字面意思首个颜色的下标为8
PaletteRecord.STANDARD_PALETTE_SIZE  常量的值为56,字面意思调色板长度为56
获取调色板第一个颜色的索引↓
```Java
short index = PaletteRecord.FIRST_COLOR_INDEX;
```

设置第一个颜色的RGB为(0,0,0)
```Java
palette.setColorAtIndex(index++, (byte) 0, (byte) 0, (byte) 0);
palette.setColorAtIndex(颜色下标, (byte)R, (byte)G, (byte)B);
```

创建颜色样式↓
```JavaJava
HSSFCellStyle style = wb.createCellStyle();
```

填充前景填充颜色↓
```Java
style.setFillForegroundColor(color.getIndex());
```

指定单元格的填充信息模式和纯色填充单元↓
```Java
style.setFillPattern(CellStyle.SOLID_FOREGROUND);
```

设置单元格边框颜色和边框宽度↓
```Java
style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框宽度
style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); //下边框颜色
style.setBorderLeft(HSSFCellStyle.BORDER_THIN); //左边框宽度
style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); //左边框颜色
style.setBorderRight(HSSFCellStyle.BORDER_THIN);  //右边框宽度
style.setRightBorderColor(IndexedColors.BLACK.getIndex());  //右边框颜色
style.setBorderTop(HSSFCellStyle.BORDER_THIN);  //上边框宽度
style.setTopBorderColor(IndexedColors.BLACK.getIndex());  //上边框颜色
```

设置单元格值居中↓
```Java
style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
```

设置字体样式↓
```Java
HSSFFont font = wb.createFont();  //创建字体
font.setFontName("宋体"); //设置字体为"宋体"
font.setFontHeightInPoints((short) 12); //设置字号
font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD); //设置字宽度
style.setFont(font);  //样式设置字体
color.put("style", style);  //将样式put进map里
```
