package co.fatboa;


import co.fatboa.Enums.MyWorkBookType;
import co.fatboa.mypoi.MySheet;
import co.fatboa.mypoi.MyWorkBook;
import org.apache.log4j.lf5.util.ResourceUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.jupiter.api.Test;

class QrCodeUtilTest {

    @Test
    void ss() {
        QrCodeUtil q = new QrCodeUtil();
        q.create("你好", "qrcode", "png", "E:\\");
    }

    @Test
    void exc() {
        MyWorkBook myWorkBook = new MyWorkBook();
        myWorkBook.setWorkBookType(MyWorkBookType.HSSFWORKBOOK);
        MySheet mySheet=new MySheet();
        myWorkBook.addSheet(mySheet);
        try {
            CellStyle cellStyle= myWorkBook.createCellStyle();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}