package co.fatboa.Enums;
/**
 * @Author: hl
 * @Description: excel类型枚举
 * @Date: 17:23 2018/8/8
 * @Modified By:
 * @Version 1.0
 */
public enum MyWorkBookType {

    HSSFWORKBOOK(0, "操作excel2003以下版本(.xls)"),
    XSSFWORKBOOK(1, "操作excel2007以上版本(.xlsx)");
    private int code;
    private String message;

    MyWorkBookType(int code, String message) {
        this.code = code;
        this.message = message;
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }
}
