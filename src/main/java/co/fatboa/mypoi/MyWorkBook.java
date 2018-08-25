package co.fatboa.mypoi;

import co.fatboa.Enums.MyCellType;
import co.fatboa.Enums.MyWorkBookType;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author: hl
 * @Description: excel表生成工具类
 * @Date: 18:25 2018/8/8
 * @Modified By:
 * @Version 1.0
 */
public class MyWorkBook {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());
    private Workbook workbook;
    private List<MySheet> mySheets = new ArrayList<MySheet>();

    public List<MySheet> getMySheets() {
        return mySheets;
    }

    public void setWorkBookType(MyWorkBookType myWorkBookType) {
        if (myWorkBookType == MyWorkBookType.HSSFWORKBOOK) {
            this.workbook = new HSSFWorkbook();
        } else {
            this.workbook = new XSSFWorkbook();
        }
    }

    public void addSheet(MySheet mySheet) {
        this.mySheets.add(mySheet);
    }

    /**
     * @return 返回单元格样式类
     * @throws Exception
     */
    public CellStyle createCellStyle() throws Exception {
        if (this.workbook == null) {
            throw new Exception("未定义workbook类型");
        }

        return this.workbook.createCellStyle();
    }

    /**
     * @return 返回字体样式类
     * @throws Exception
     */
    public Font createFont() throws Exception {
        if (this.workbook == null) {
            throw new Exception("未定义workbook类型");
        }
        return this.workbook.createFont();
    }

    /**
     * 创建文本单元
     *
     * @param sheet
     * @param myCell 单元格配置信息
     */
    public void writeTextCell(Sheet sheet, MyCell myCell) {
        int rowIndex = myCell.getRowIndex();
        int colIndex = myCell.getColIndex();
        int rowspan = myCell.getRowSpan();
        int colspan = myCell.getColSpan();
        for (int i = rowIndex; i < rowIndex + rowspan; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            for (int j = colIndex; j < colIndex + colspan; j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j, CellType.STRING);
                }
            }
        }

        if (rowspan > 1 && colspan > 1) {
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + rowspan - 1, colIndex, colIndex + colspan - 1));
        }
        Row firstRow = sheet.getRow(rowIndex);
        if (myCell.getHeight() > 0) {
            firstRow.setHeight((short) myCell.getHeight());
        }
        if (myCell.getWidth() > 0) {
            sheet.setColumnWidth(colIndex, myCell.getWidth());
        }
        Cell firstCell = firstRow.getCell(colIndex);
        firstCell.setCellStyle(myCell.getCellStyle());
        firstCell.setCellValue(myCell.getText());
    }

    /**
     * 创建图片单元
     *
     * @param sheet
     * @param myCell  单元格配置信息
     * @param stretch 图片是否延伸
     */
    public void writeImageCell(Sheet sheet, MyCell myCell, boolean stretch) {
        byte[] buffer = null;
        try {
            FileInputStream inputStream = new FileInputStream(new File(myCell.getPath()));
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            byte[] b = new byte[1024];
            int n;
            while ((n = inputStream.read(b)) != -1) {
                outputStream.write(b, 0, n);
            }
            buffer = outputStream.toByteArray();
            inputStream.close();
            outputStream.close();
        } catch (IOException e) {
            logger.error("读取图片数据失败:" + e.getMessage());
            e.printStackTrace();
        }
        int colIndex1 = myCell.getColIndex();
        int rowIndex1 = myCell.getRowIndex();
        int colIndex2 = myCell.getColIndex() + myCell.getColSpan() - 1;
        int rowIndex2 = myCell.getRowIndex() + myCell.getRowSpan() - 1;
        int picIndex = this.workbook.addPicture(buffer, Workbook.PICTURE_TYPE_PNG);
        HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor(50, 50, 0, 0, (short) colIndex1, rowIndex1, (short) colIndex2, rowIndex2);
        HSSFPicture picture = patriarch.createPicture(anchor, picIndex);
        if (stretch) {
            picture.resize();
        }
    }

    /**
     * 创建所有单元格
     *
     * @param sheet
     * @param myCells
     */
    public void writeCells(Sheet sheet, List<MyCell> myCells) {
        for (MyCell myCell : myCells) {
            if (myCell.getType() == MyCellType.Text) {
                writeTextCell(sheet, myCell);
            } else if (myCell.getType() == MyCellType.Image) {
                writeImageCell(sheet, myCell, false);
            }
        }
    }

    /**
     * 创建多个sheet
     */
    private void writeSheet() {
        for (MySheet mySheet : mySheets) {
            Sheet sheet = this.workbook.createSheet();
            writeCells(sheet, mySheet.getMyCells());
        }
    }

    /**
     * @return workbook数据
     */
    public Workbook createExcel() {
        writeSheet();
        return this.workbook;
    }

    /**
     * excel表数据转换为jsonArray
     * 确保excel表中每个sheet第一行为属性名,且结构均相同
     * 表内不可有图片
     *
     * @param inputStream  excel输入流
     * @param workBookType excel表版本
     * @return 返回JsonArray
     */

    public JsonArray importExcel(InputStream inputStream, MyWorkBookType workBookType) {
        Workbook workbook = null;
        try {
            if (workBookType == MyWorkBookType.HSSFWORKBOOK) {
                workbook = new HSSFWorkbook(inputStream);
            } else if (workBookType == MyWorkBookType.XSSFWORKBOOK) {
                workbook = new XSSFWorkbook(inputStream);
            }
        } catch (IOException e) {
            logger.error("创建workbook失败:", e.getMessage());
            e.printStackTrace();
        }

        JsonArray jsonArray = new JsonArray(); // 存放excel表中每一行数据

        for (int sheetnum = 0; sheetnum < workbook.getNumberOfSheets(); sheetnum++) {
            Sheet sheet = workbook.getSheetAt(sheetnum);
            if (sheet == null) {
                continue;
            }

            Row firstRow = sheet.getRow(0);
            if (firstRow == null) {
                logger.error("第一行数据属性名,不能为空!");
                break;
            }
            List<String> cellnames = new ArrayList<String>(); //存放属性名数组
            // 一般第一行存放数据属性名
            for (short cellnum = 0; cellnum <= firstRow.getLastCellNum(); cellnum++) {
                Cell cell = firstRow.getCell(cellnum);
                cellnames.add(getValue(cell));
            }

            for (int rownum = 1; rownum < sheet.getLastRowNum(); rownum++) {
                Row row = sheet.getRow(rownum);
                if (row == null) {
                    continue;
                }
                JsonObject object = new JsonObject();
                if (cellnames != null && cellnames.size() > 0) {
                    for (int i = 0; i < cellnames.size(); i++) {
                        Cell cell = row.getCell(i);
                        if (cell == null) {
                            continue;
                        }
                        object.addProperty(cellnames.get(i), getValue(cell));
                    }
                }
                jsonArray.add(object);
            }

        }
        return jsonArray;
    }

    /**
     * 将单元格值类型为字符串型
     *
     * @param cell
     * @return
     */
    public String getValue(Cell cell) {
        DecimalFormat df = new DecimalFormat("###################.###########");
        if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(df.format(cell.getNumericCellValue()));
        } else if (cell.getCellType() == cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }

    /**
     * 读取excel中图片
     * @param sheet 注意传XSSFSheet类型
     * @return
     */
    public Map<String, PictureData> getPictureDataMap(XSSFSheet sheet) {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            XSSFDrawing drawing = (XSSFDrawing) dr;
            List<XSSFShape> shapeList = drawing.getShapes();
            if (shapeList != null && shapeList.size() > 0) {
                for (XSSFShape shape : shapeList) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = ctMarker.getRow() + "";
                    map.put(picIndex, pic.getPictureData());
                }
            }
        }
        return map;
    }
}
