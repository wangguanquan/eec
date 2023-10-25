/*
 * Copyright (c) 2017-2018, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ttzero.excel.entity;

import org.ttzero.excel.util.FileSignatures;

import javax.imageio.ImageIO;
import java.awt.Graphics2D;
import java.awt.Color;
import java.awt.BasicStroke;
import java.awt.AlphaComposite;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

/**
 * 水印，Excel并没有水印功能，EEC的水印功能是使用Excel的背景实现，打印的时候该背景会被忽略。
 *
 * <p>注意：只接受{@link FileSignatures#whitelist}白名单的图片格式</p>
 *
 * @author guanquan.wang at 2018-01-26 15:23
 */
public class WaterMark {
    /**
     * 水印图片临时路径
     */
    private final Path imagePath;
    /**
     * 标记是否为临时创建，文本和InputStream流时为true
     */
    private boolean temp;
    /**
     * 水印图片签名
     */
    private final FileSignatures.Signature signature;
    /**
     * 使用一段文本创建水印
     *
     * @param word 一段有意义的文本
     */
    public WaterMark(String word) {
        imagePath = createWaterMark(word);
        // 水印由内部制作所以这里不需要检查签名
        signature = new FileSignatures.Signature("png", FileSignatures.whitelist.getOrDefault("png", "image/png"), 510, 300);
    }

    /**
     * 使用本地图片创建水印
     *
     * @param imagePath 图片的本地文件
     */
    public WaterMark(Path imagePath) {
        this.imagePath = imagePath;
        signature = FileSignatures.test(imagePath);
    }

    /**
     * 使用图片流创建水印，可以下载远程图片创建水印
     *
     * @param inputStream 图片流
     * @throws IOException 读取流异常
     */
    public WaterMark(InputStream inputStream) throws IOException {
        imagePath = createTemp();
        Files.copy(inputStream, imagePath, StandardCopyOption.REPLACE_EXISTING);
        signature = FileSignatures.test(imagePath);
    }

    /**
     * 获取水印图片路径
     *
     * @return 水印图片临时路径，
     */
    public Path get() {
        // 非白名单图片格式返回null
        return canWrite() ? imagePath : null;
    }

    /**
     * 使用一段文本创建水印
     *
     * @param mark 文本
     * @return 水印对象
     */
    public static WaterMark of(String mark) {
        return new WaterMark(mark);
    }

    /**
     * 使用本地图片创建水印
     *
     * @param path 图片的本地文件
     * @return 水印对象WaterMark
     */
    public static WaterMark of(Path path) {
        return new WaterMark(path);
    }

    /**
     * 使用图片流创建水印，可以下载远程图片创建水印
     *
     * @param is 图片流
     * @return 水印对象WaterMark
     * @throws IOException 读取流异常
     */
    public static WaterMark of(InputStream is) throws IOException {
        return new WaterMark(is);
    }

    /**
     * 使用指定文本制作一张水印图片
     *
     * @param watermark 一段有意义的文本
     * @return the temp image path
     */
    private Path createWaterMark(String watermark) {
        try {
            Path temp = createTemp();
            int width = 510;
            int height = 300;

            BufferedImage bi = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            // Setting white background color
            int minx = bi.getMinX();
            int miny = bi.getMinY();
            for (int i = minx; i < width; i++) {
                for (int j = miny; j < height; j++) {
                    bi.setRGB(i, j, 0xffffff);
                }
            }

            Graphics2D g2d = bi.createGraphics();
            g2d.setColor(new Color(200, 200, 200));
            g2d.setStroke(new BasicStroke(1));
            g2d.setFont(new java.awt.Font("华文细黑", java.awt.Font.ITALIC, 50));
            g2d.rotate(Math.toRadians(-10));

            for (int i = 1; i < 10; i++) {
                g2d.drawString(watermark, 0, 60 * i);
            }
            // Setting alpha
            g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
            g2d.dispose();
            ImageIO.write(bi, "png", temp.toFile());
            return temp;
        } catch (IOException e) {
            throw new ExcelWriteException("创建水印失败.", e);
        }
    }

    /**
     * 创建临时文件
     *
     * @return 临时文件路径
     * @throws IOException 没有权限或者磁盘不足等情况
     */
    private Path createTemp() throws IOException {
        temp = true;
        return Files.createTempFile("waterMark", "png");
    }

    /**
     * 删除临时文件，传入InputStream或文本时会保存到临时文件，所以需要清理资源
     *
     * @return 出异常时返回false
     */
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

    /**
     * 获取水印图片的后缀
     *
     * @return 水印图片后缀
     */
    public String getSuffix() {
        return "." + signature.extension;
    }

    /**
     * 水印图片Content-type
     *
     * @return Content-type
     */
    public String getContentType() {
        return signature.contentType;
    }

    /**
     * 测试水印图片是否可输出
     *
     * @return true: 资源可信任输出到Excel
     */
    public boolean canWrite() {
        return signature.isTrusted();
    }
}
