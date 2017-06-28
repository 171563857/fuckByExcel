poi操作
=======

创建一个excel关联对象HSSFWorkbook：
```
HSSFWorkbook book = new HSSFWorkbook();
```

创建一个sheet：
```
HSSFSheet st = book.createSheet("sheet1");
```     

创建第i行：
```
HSSFRow row = st.createRow(i);
```

创建第i行的j列：
```
HSSFCell cell = row.createCell(j);
```

设置cell属性
给单元格设置边框属性：
```
HSSFCellStyle style = book.createCellStyle();
// 左右上下边框样式
style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
style.setBorderRight(HSSFCellStyle.BORDER_THIN);
style.setBorderTop(HSSFCellStyle.BORDER_THIN);
style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
// 左右上下边框颜色(深蓝色),只有设置了上下左右边框以后给边框设置颜色才会生效
style.setLeftBorderColor(HSSFColor.BLACK.index);
style.setRightBorderColor(HSSFColor.BLACK.index);
style.setTopBorderColor(HSSFColor.BLACK.index);
style.setBottomBorderColor(HSSFColor.BLACK.index);
```
    
给单元格设置背景：
```
style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);// 设置了背景色才有效果
style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
    
给单元格设置字体：
```
// 单元格字体
HSSFFont font = book.createFont();
font.setFontName("宋体");
```

设置字体以后，需要把字体加入到style中：
```
style.setFont(font);
```

设置好单元格属性以后，需要这种属性的单元格就可以调用此style：
```
cell.setCellStyle(style);
```

设置sheet表单的列宽：
```
st.setColumnWidth(i, cellWidths.get(i).intValue() * 160);
```

列宽的设置方法在HSSFSheet中，方法参数：第一个参数表示第几列，从0开始数；第二个参数表示宽度为多少，大小由使用者调整。

合并单元格：
```
st.addMergedRegion(new CellRangeAddress(0, 1, 0, keys.size() - 1));
```

单元格合并方法也是在HSSFSheet中，
方法参数：一个CellRangeAddress，
该类构造函数的4个参数分别表示为：合并开始行，合并结束行，合并开始列，合并结束列

注：合并方法最好写在最后面，不然有可能会影响到某些单元格添加单元格属性的操作

下面是我写的一个根据传入的数据，把数据导出到excel的接口：
```
/**
 * 导出到excel 导出的路径以及导出文件名称在配置文件中定义
 * 在任何地方此方法可以作为组件调用的，只需要提供需要保存的数据，每一列的属性，以及对应的中文名称，每一列的宽度，文件路径，文件名称
 * 
 * @param list
 *            the data which will be saved to excel
 * @param keys
 *            the key of the column
 * @param cnames
 *            the name described in Chinese
 * @param cellWidths
 *            the width of olumns
 * @param excelPath
 *            the path of excel
 * @param fileName
 *            the name of the file in server
 */
public HSSFWorkbook doExportResults(List<Map<String, Object>> list,
    List<String> keys, List<String> cnames, List<Integer> cellWidths,
    String excelPath, String fileName) {

  File excel = new File(excelPath);// 创建文件
  String sheetName = fileName.substring(0, fileName.lastIndexOf("."));
  String dateStr = fileName.substring(fileName.indexOf("_") + 1,
      fileName.lastIndexOf("."));

  HSSFWorkbook book = new HSSFWorkbook();// 创建excel
  HSSFSheet st = book.createSheet(sheetName);
  // 第一行，标题
  HSSFRow row = st.createRow(0);
  HSSFCell cell = row.createCell(0);
  cell.setCellValue(sheetName);
  // 单元格属性 第一行的属性在最后设置
  HSSFCellStyle style = book.createCellStyle();
  // 左右上下边框样式
  style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
  style.setBorderRight(HSSFCellStyle.BORDER_THIN);
  style.setBorderTop(HSSFCellStyle.BORDER_THIN);
  style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
  // 左右上下边框颜色(深蓝色)
  style.setLeftBorderColor(HSSFColor.BLACK.index);
  style.setRightBorderColor(HSSFColor.BLACK.index);
  style.setTopBorderColor(HSSFColor.BLACK.index);
  style.setBottomBorderColor(HSSFColor.BLACK.index);
  // 背景
  style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);// 设置了背景色才有效果
  style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);

  // 单元格字体
  HSSFFont font = book.createFont();
  font.setFontName("宋体");
  style.setFont(font);

  // 第二行，日期
  row = st.createRow(2);
  cell = row.createCell(0);
  cell.setCellValue("导出日期");
  cell = row.createCell(1);
  cell.setCellValue(dateStr);
  //为日期行设置单元格属性
  for (int i = 0; i < keys.size(); i++) {
    if (row.getCell(i) != null) {
      cell = row.getCell(i);
    } else {
      cell = row.createCell(i);
    }
    cell.setCellStyle(style);
  }
  // 第三行，表头
  row = st.createRow(3);
  for (int i = 0; i < keys.size(); i++) {
    cell = row.createCell(i);
    cell.setCellValue(cnames.get(i));
    cell.setCellStyle(style);
    st.setColumnWidth(i, cellWidths.get(i).intValue() * 160);
  }
  for (int i = 0; i < list.size(); i++) {// 创建每一行数据
    row = st.createRow(4 + i);
    Map<String, Object> tmp = list.get(i);
    if (tmp == null || tmp.isEmpty()) {
      continue;
    }
    for (int j = 0; j < keys.size(); j++) {// 设置每一行的每一个单元格的值
      cell = row.createCell(j);
      cell.setCellStyle(style);
      Object obj = tmp.get(keys.get(j));
      if (obj == null) {
        cell.setCellValue("");
      } else {
        cell.setCellValue(obj.toString());
      }
    }
  }
  // 合并单元格
  HSSFCellStyle s1 = book.createCellStyle();
  s1.cloneStyleFrom(style);
  s1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
  s1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
  s1.setWrapText(true);
  s1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);// 设置了背景色才有效果
  s1.setFillForegroundColor(HSSFColor.BROWN.index);
  HSSFFont fo = book.createFont();

  fo.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
  fo.setFontHeight((short) 350);
  fo.setFontName("宋体");
  s1.setFont(fo);
  st.getRow(0).getCell(0).setCellStyle(s1);
  st.addMergedRegion(new CellRangeAddress(0, 1, 0, keys.size() - 1));
  st.addMergedRegion(new CellRangeAddress(2, 2, 1, keys.size() - 1));
  // 创建表格结束
  FileOutputStream out = null;
  try {
    if (!excel.getParentFile().exists()) {
      System.out.println(excel.getParentFile().mkdirs());
    }
    if (!excel.exists()) {
      System.out.println(excel.createNewFile());
    }
    out = new FileOutputStream(excel);
    book.write(out);// 把excel写入到本地文件
  } catch (FileNotFoundException e) {
    logger.error("导出到文件时找不到文件:" + e.getMessage());
  } catch (IOException e) {
    logger.error("导出到文件时输出流错误:" + e.getMessage());
  } finally {
    try {
      if (out != null) {
        out.close();
      }
    } catch (IOException e) {
      logger.error("导出到文件时关闭输出流错误:" + e.getMessage());
    }
  }
  return book;
}
```  
原文转摘自:http://www.cnblogs.com/God-froest/p/excel_1.html
