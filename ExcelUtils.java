package fuckExcel;

import fuckExcel.CellDTO;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.Map;

/**
 * author: GoL
 * time:   2017-05-12
 */
public class Excel2003Utils {
    /**
     * 生成填充前景颜色样式
     * @param color 颜色下标
     */
    public static HSSFCellStyle createStyle(HSSFWorkbook wb, HSSFColor color) {
        HSSFCellStyle style = wb.createCellStyle();
        //设置颜色
        if (null != color) {
            style.setFillForegroundColor(color.getIndex());
            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        }
        //设置边框
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        //设置居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置字体
        HSSFFont font = wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(font);
        return style;
    }

    /**
     * 生成超链接样式
     * @param color 颜色下标
     */
    public static HSSFCellStyle createLinkStyle(HSSFWorkbook wb, short color) {
        HSSFCellStyle  style= wb.createCellStyle();
        HSSFFont cellFont= wb.createFont();
        cellFont.setUnderline((byte) 1);
        cellFont.setColor(color);
        style.setFont(cellFont);
        return style;
    }

    /**
     * 生成字体样式
     * @param color 颜色下标
     */
    public static HSSFCellStyle createFontStyle(HSSFWorkbook wb, short color) {
        HSSFCellStyle style = wb.createCellStyle();
        //设置边框
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        //设置居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置字体
        HSSFFont font = wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setColor(color);
        style.setFont(font);
        return style;
    }



    /**
     * 合并单元格
     */
    public static void mergeCells(Sheet sheet, int firstRow, int lastRow, int firstColumn, int lastColumn) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
    }
    
    /**
     * 生成表头标题
     */
    public static void createTitle(Sheet sheet, Row row, List<CellDTO> titles) {
        Cell cell;
        int index = row.getPhysicalNumberOfCells();
        for (CellDTO dto : titles) {
            cell = row.createCell(index++);
            adjustColumnWidth(sheet, dto.getTitle(), index - 1);
            cell.setCellValue(dto.getTitle());
            cell.setCellStyle(dto.getStyle());
        }
    }

    /**
     * 生成表头分组
     */
    public static void createGroup(Sheet sheet, Row row, List<CellDTO> cellDTOs) {
        Cell cell;
        int index = row.getPhysicalNumberOfCells();
        for (CellDTO dto : cellDTOs) {
            int len = index + dto.getWidth();
            Excel2003Utils.mergeCells(sheet, row.getRowNum(), row.getRowNum(), index, len - 1);
            adjustColumnWidth(sheet, dto.getTitle(), index);
            cell = row.createCell(index++);
            cell.setCellValue(dto.getTitle());
            cell.setCellStyle(dto.getStyle());
            while (index < len) {
                cell = row.createCell(index++);
                cell.setCellStyle(dto.getStyle());
            }
        }
    }

    /**
     * 填充表格数据
     * @param title 标题
     * @param dataList  数据
     * @param index 下标,从第几行开始填充数据
     */
    public static void createBody(Sheet sheet, List<CellDTO> title, List<Map<String, String>> dataList, int index) {
        Row row;
        Cell cell;
        for (Map<String, String> data : dataList) {
            row = sheet.createRow(index++);
            int index2 = 0;
            for (CellDTO c : title) {
                cell = row.createCell(index2++);
                cell.setCellValue(data.get(c.getName()));
            }
        }
    }
    /**
     * 设置列宽
     */
    public static void adjustColumnWidth(Sheet sheet, String value, int index) {
        int columnWidth = value.length() * 700 > sheet.getColumnWidth(index) ?value.length() * 700 : sheet.getColumnWidth(index);
        sheet.setColumnWidth(index, columnWidth);
    }
}

