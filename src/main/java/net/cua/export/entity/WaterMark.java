package net.cua.export.entity;

import net.cua.export.manager.Const;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.nio.file.*;

/**
 * Created by guanquan.wang at 2018-01-26 15:23
 */
public class WaterMark {
    private Logger logger = LogManager.getLogger(getClass());
    private Path imagePath;
    private boolean temp;

    public WaterMark(String word) { // 文字水印
        imagePath = createWaterMark(word);
    }

    public WaterMark(Path imagePath) {  // 图片水印（路径）
        this.imagePath = imagePath;
    }

    public WaterMark(InputStream inputStream) throws IOException { // 图片水印（流）
        imagePath = createTemp();
        Files.copy(inputStream, imagePath, StandardCopyOption.REPLACE_EXISTING);
    }

    public Path get() {
        return imagePath;
    }

    public static WaterMark of(String mark) {
        return new WaterMark(mark);
    }

    public static WaterMark of(Path path) {
        return new WaterMark(path);
    }

    public static WaterMark of(InputStream is) throws IOException {
        return new WaterMark(is);
    }

    /**
     * 生成水印图片
     *
     * @param watermark
     * @return
     * @throws IOException
     */
    public Path createWaterMark(String watermark) {
        try {
            Path temp = createTemp();
            int width = 510; // 水印图片的宽度
            int height = 300; // 水印图片的高度 因为设置其他的高度会有黑线，所以拉高高度

            // 获取bufferedImage对象
            BufferedImage bi = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            // 处理背景色，设置为 白色
            int minx = bi.getMinX();
            int miny = bi.getMinY();
            for (int i = minx; i < width; i++) {
                for (int j = miny; j < height; j++) {
                    bi.setRGB(i, j, 0xffffff);
                }
            }

            // 获取Graphics2d对象
            Graphics2D g2d = bi.createGraphics();
            // 设置字体颜色为灰色
            g2d.setColor(new Color(200, 200, 200));
            // 设置图片的属性
            g2d.setStroke(new BasicStroke(1));
            // 设置字体
            g2d.setFont(new java.awt.Font("华文细黑", java.awt.Font.ITALIC, 50));
            // 设置字体倾斜度
            g2d.rotate(Math.toRadians(-10));

            // 写入水印文字 原定高度过小，所以累计写水印，增加高度
            for (int i = 1; i < 10; i++) {
                g2d.drawString(watermark, 0, 60 * i);
            }
            // 设置透明度
            g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
            // 释放对象
            g2d.dispose();
            ImageIO.write(bi, "png", temp.toFile());
            return temp;
        } catch (IOException e) {
            logger.error("", e);
        }
        return null;
    }

    private Path createTemp() throws IOException {
        temp = true;
        return Files.createTempFile("waterMark", "png");
    }

    public boolean delete() {
        if (imagePath != null && temp) {
            try {
                Files.deleteIfExists(imagePath);
            } catch (IOException e) {
                return false;
            }
        }
        return true;
    }

    public String getSuffix() {
        String suffix = null;
        if (temp) {
            suffix = Const.Suffix.PNG;
        } else if (imagePath != null) {
            String name = imagePath.getFileName().toString();
            int n;
            if ((n = name.lastIndexOf('.')) > 0) {
                suffix = name.substring(n);
            }
        }
        return suffix != null ? suffix : Const.Suffix.PNG;
    }

    public String getContentType() {
        String suffix = getSuffix().substring(1).toUpperCase();
        Field[] fields = Const.ContentType.class.getDeclaredFields();
        for (Field f : fields) {
            if (f.getName().equals(suffix)) {
                try {
                    return f.get(null).toString();
                } catch (IllegalAccessException e) {
                    ; // Empty
                }
            }
        }
        return Const.ContentType.PNG;
    }
}
