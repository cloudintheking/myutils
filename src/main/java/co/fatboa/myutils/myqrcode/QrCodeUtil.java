package co.fatboa.myutils.myqrcode;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.EnumMap;

/**
 * @Author: hl
 * @Description: 二维码生成工具类
 * @Date: 10:31 2018/8/8
 * @Modified By:
 * @Version 1.0
 */
public class QrCodeUtil {
    private static final Logger logger = LoggerFactory.getLogger(QrCodeUtil.class);
    private static int BLACK = 0xFF000000;//二维码颜色
    private static int WHITE = 0xFFFFFFFF;//二维码背景颜色
    private static int width = 300;//二维码图片宽度
    private static int height = 300;//二维码图片高度

    //二维码格式参数
    private static EnumMap<EncodeHintType, Object> hints = new EnumMap<EncodeHintType, Object>(EncodeHintType.class);

    static {
        /*二维码的纠错级别(排错率),4个级别：
         L (7%)、
         M (15%)、
         Q (25%)、
         H (30%)(最高H)
         纠错信息同样存储在二维码中，纠错级别越高，纠错信息占用的空间越多，那么能存储的有用讯息就越少；共有四级；
         选择M，扫描速度快。
         */
        hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
        // 二维码边界空白大小 1,2,3,4 (4为默认,最大)
        hints.put(EncodeHintType.MARGIN, 1);
        hints.put(EncodeHintType.CHARACTER_SET, "UTF-8");
        hints.put(EncodeHintType.MAX_SIZE, 350);
        hints.put(EncodeHintType.MARGIN.MIN_SIZE, 150);
    }

    public QrCodeUtil(int BLACK, int WHITE, int width, int height) {
        this.BLACK = BLACK;
        this.WHITE = WHITE;
        this.width = width;
        this.height = height;
    }

    public QrCodeUtil() {
    }

    /**
     * 绘制二维码(无logo)
     *
     * @param contents 二维码内容
     * @return 二维码图片
     */
    public static BufferedImage encodeImage(String contents) {
        BufferedImage image = null;
        try {
            BitMatrix matrix = new MultiFormatWriter().encode(contents, BarcodeFormat.QR_CODE, width, height, hints);
            image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            int width = matrix.getWidth();
            int height = matrix.getHeight();
            for (int x = 0; x < width; x++) {
                for (int y = 0; y < height; y++) {
                    image.setRGB(x, y, matrix.get(x, y) ? BLACK : WHITE);
                }
            }
        } catch (WriterException e) {
            logger.error("生成二维码失败" + e.getMessage());
            e.printStackTrace();
        }
        return image;
    }

    /**
     * 绘制logo到二维码图中
     *
     * @param qrImg    二维码图
     * @param logoPath
     * @return 带logo的二维码图片
     */
    public static BufferedImage encodeImgLogo(BufferedImage qrImg, String logoPath) {
        try {
            Graphics2D g = qrImg.createGraphics();
            File logoImg = new File(logoPath);
            if (!logoImg.exists()) {
                throw new Exception("文件不存在");
            }
            BufferedImage logo = ImageIO.read(logoImg);
            //设置二维码大小，太大，会覆盖二维码，此处30%
            int logoWidth = logo.getWidth(null) > qrImg.getWidth() * 3 / 10 ? (qrImg.getWidth() * 3 / 10) : logo.getWidth(null);
            int logoHeight = logo.getHeight(null) > qrImg.getWidth() * 3 / 10 ? (qrImg.getHeight() * 3 / 10) : logo.getHeight(null);

            //设置logo图片放置位置
            //中心
            int x = (qrImg.getWidth() - logoWidth) / 2;
            int y = (qrImg.getHeight() - logoHeight) / 2;
            //开始合并绘制图片
            g.drawImage(logo, x, y, logoWidth, logoHeight, null);
            g.drawRoundRect(x, y, logoWidth, logoHeight, 15, 15);
            //logo边框大小
            g.setStroke(new BasicStroke(2));
            //logo边框颜色
            g.setColor(Color.WHITE);
            g.drawRect(x, y, logoWidth, logoHeight);
            g.dispose();
            logo.flush();
            qrImg.flush();
        } catch (Exception e) {
            logger.error(e.getMessage());
            e.printStackTrace();
        }

        return qrImg;
    }

    /**
     * @param contents 二维码内容
     * @param qrName   生成二维码图片名
     * @param format   生成二维码图片格式
     * @param outPath  生成二维码图片目录路径
     */
    public static void create(String contents, String qrName, String format, String outPath) {
        BufferedImage image = encodeImage(contents);
        File outDir = new File(outPath);
        if (!outDir.exists()) {
            outDir.mkdirs();
        }
        String qrPath = outPath + qrName + "." + format;
        File qrImg = new File(qrPath);
        try {
            ImageIO.write(image, format, qrImg);
            logger.info("生成成功:" + qrPath);
        } catch (IOException e) {
            logger.error("二维码写入文件失败:" + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * @param contents 二维码内容
     * @param qrName   生成二维码图片名
     * @param format   生成二维码图片格式
     * @param logoPath logo路径
     * @param outPath  生成二维码图片目录路径
     */
    public void create(String contents, String qrName, String format, String logoPath, String outPath) {
        BufferedImage image = encodeImage(contents);
        image = encodeImgLogo(image, logoPath);
        File outDir = new File(outPath);
        if (!outDir.exists()) {
            outDir.mkdirs();
        }
        String qrPath = outPath + qrName + "." + format;
        File qrImg = new File(qrPath);
        try {
            ImageIO.write(image, format, qrImg);
            logger.info("生成成功:" + qrPath);
        } catch (IOException e) {
            logger.error("二维码写入文件失败:" + e.getMessage());
            e.printStackTrace();
        }
    }

}
