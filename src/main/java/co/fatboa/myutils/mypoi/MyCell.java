package co.fatboa.myutils.mypoi;

import co.fatboa.myutils.mypoi.enums.MyCellType;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @Author: hl
 * @Description: cell参数配置类
 * @Date: 16:49 2018/8/8
 * @Modified By:
 * @Version 1.0
 */
public class MyCell {
    private MyCellType type; //类型
    private String text; //文本内容
    private int colIndex; //所在列下标(0,1,2,...)
    private int rowIndex; //所在行下标(0,1,2,...)
    private int colSpan; //占据列数(1,2,3,...)
    private int rowSpan; //占据行数(1,2,3,...)
    private int height; //高度 可不配置,为默认高度
    private int width; //宽度 可不配置,为默认宽度
    private CellStyle cellStyle; //单元格样式
    private String path; // 图片路径

    public MyCell(MyCellType type, String text, int colIndex, int rowIndex, int colSpan, int rowSpan, int height, int width, CellStyle cellStyle, String path) {
        this.type = type;
        this.text = text;
        this.colIndex = colIndex;
        this.rowIndex = rowIndex;
        this.colSpan = colSpan;
        this.rowSpan = rowSpan;
        this.height = height;
        this.width = width;
        this.cellStyle = cellStyle;
        this.path = path;
    }

    public MyCell() {
    }

    public MyCellType getType() {
        return type;
    }

    public void setType(MyCellType type) {
        this.type = type;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void setColIndex(int colIndex) {
        this.colIndex = colIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getColSpan() {
        return colSpan;
    }

    public void setColSpan(int colSpan) {
        this.colSpan = colSpan;
    }

    public int getRowSpan() {
        return rowSpan;
    }

    public void setRowSpan(int rowSpan) {
        this.rowSpan = rowSpan;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }
}
