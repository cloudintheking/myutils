package co.fatboa.myutils;


import co.fatboa.myutils.mypoi.MyCell;
import co.fatboa.myutils.mypoi.enums.MyCellType;
import co.fatboa.myutils.mypoi.enums.MyWorkBookType;
import co.fatboa.myutils.mypoi.MySheet;
import co.fatboa.myutils.mypoi.MyWorkBook;
import co.fatboa.myutils.myqrcode.QrCodeUtil;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;

class QrCodeUtilTest {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    /**
     * 生成二维码
     */
    @Test
    void ss() {
        QrCodeUtil.create("你好", "qrcode", "png", "I:\\");
    }

    /**
     * 封装poi测试
     */
    @Test
    void exc() {
        MyWorkBook myWorkBook = new MyWorkBook();//实例一个MyWorkBook
        myWorkBook.setWorkBookType(MyWorkBookType.HSSFWORKBOOK);//设置hssf格式
        MySheet mySheet = new MySheet();//实例一个MySheet
        try {
            CellStyle cellStyle = myWorkBook.createCellStyle();//获取单元格样式
            cellStyle.setAlignment(HorizontalAlignment.CENTER); //水平居中
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
            cellStyle.setWrapText(true); //自动换行
            Font font = myWorkBook.createFont();//获取字体样式
            font.setBold(true);//字体加粗
            cellStyle.setFont(font);//设置单元格的字体样式
            MyCell myCellc1r1 = new MyCell(MyCellType.Text, "姓名", 0, 0, 1, 1, 0, 0, cellStyle, null);//实例一个mycell
            MyCell myCellc1r2 = new MyCell(MyCellType.Text, "性别", 0, 1, 1, 1, 0, 0, cellStyle, null);
            MyCell myCellc1r3 = new MyCell(MyCellType.Text, "姓年龄", 0, 2, 1, 1, 0, 0, cellStyle, null);
            MyCell myCellc2r1 = new MyCell(MyCellType.Text, "张三", 1, 0, 1, 1, 0, 0, cellStyle, null);
            MyCell myCellc2r2 = new MyCell(MyCellType.Text, "男", 1, 1, 1, 1, 0, 0, cellStyle, null);
            MyCell myCellc2r3 = new MyCell(MyCellType.Text, "25", 1, 2, 1, 1, 0, 0, cellStyle, null);
            MyCell myCellc3r1 = new MyCell(MyCellType.Image, null, 2, 0, 2, 3, 0, 0, cellStyle, "src/main/resources/static/img/user.jpg");
            mySheet.add(myCellc1r1);//将单元信息加入sheet
            mySheet.add(myCellc1r2);
            mySheet.add(myCellc1r3);
            mySheet.add(myCellc2r1);
            mySheet.add(myCellc2r2);
            mySheet.add(myCellc2r3);
            mySheet.add(myCellc3r1);
            Workbook workbook = myWorkBook.addSheet(mySheet).createExcel();//将mysheet加入myworkbook,并生成workbook
            FileOutputStream fileOutputStream = new FileOutputStream(new File("I:/temp.xls"));//到处位置
            workbook.write(fileOutputStream);//数据写入文件
            logger.info("excel自定义表生成成功");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}