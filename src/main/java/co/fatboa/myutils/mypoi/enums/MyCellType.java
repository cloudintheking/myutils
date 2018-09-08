package co.fatboa.myutils.mypoi.enums;

/**
 * @Author: hl
 * @Description: 单元格类型枚举
 * @Date: 16:49 2018/8/8
 * @Modified By:
 * @Version 1.0
 */
public enum MyCellType {
    Text(0, "文本"),
    Image(1, "图片");

    private int code;
    private String message;

    MyCellType(int code, String message) {
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
