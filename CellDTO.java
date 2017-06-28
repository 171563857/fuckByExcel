package com.evada.de.designtable.excelPRO.Model;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

/**
 * author: GoL
 * time:   2017-06-15
 * Cell单元格扩展类
 */
public class CellDTO {
    private String name;            //英文标识
    private String title;           //中文标识
    private int width = 1;          //长度
    private int height = 1;         //高度
    private HSSFCellStyle style;    //样式

    public CellDTO() {
    }

    public CellDTO(String name, HSSFCellStyle style){
        this.name = name;
        this.title = name;
        this.style = style;
    }

    public CellDTO(String name, String title, HSSFCellStyle style) {
        this.name = name;
        this.title = title;
        this.style = style;
    }

    public CellDTO(String name, int width, HSSFCellStyle style) {
        this.name = name;
        this.title = name;
        this.width = width;
        this.style = style;
    }

    public CellDTO(String name, String title, int width, HSSFCellStyle style) {
        this.name = name;
        this.title = title;
        this.width = width;
        this.style = style;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public HSSFCellStyle getStyle() {
        return style;
    }

    public void setStyle(HSSFCellStyle style) {
        this.style = style;
    }
}
