/*
 * Copyright (c) 2017-2018, guanquan.wang@hotmail.com All Rights Reserved.
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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.util.FileSignatures;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

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
 * 水印，Excel并没有水印功能，本工具的水印功能是使用Excel的背景实现，打印的时候该背景会被忽略。
 *
 * <p>注意：只接受{@link FileSignatures#whitelist}白名单的图片格式</p>
 *
 * @author guanquan.wang at 2018-01-26 15:23
 */
public class Watermark {
    /**
     * LOGGER
     */
    private final static Logger LOGGER = LoggerFactory.getLogger(Watermark.class);
    /**
     * 水印图片临时路径
     */
    private Path imagePath;
    /**
     * 标记是否为临时创建，文本和InputStream流时为true
     */
    private boolean temp;
    /**
     * 水印图片签名
     */
    private FileSignatures.Signature signature;
    /**
     * 水印文本
     */
    private String txt;
    /**
     * 使用一段文本创建水印
     *
     * @param word 一段有意义的文本
     */
    public Watermark(String word) {
        this.txt = word;
    }

    /**
     * 使用本地图片创建水印
     *
     * @param imagePath 图片的本地文件
     */
    public Watermark(Path imagePath) {
        this.imagePath = imagePath;
    }

    /**
     * 使用图片流创建水印，可以下载远程图片创建水印
     *
     * @param inputStream 图片流
     * @throws IOException 读取流异常
     */
    public Watermark(InputStream inputStream) throws IOException {
        imagePath = createTemp();
        Files.copy(inputStream, imagePath, StandardCopyOption.REPLACE_EXISTING);
    }

    /**
     * 获取水印图片路径
     *
     * @return 水印图片临时路径，
     */
    public Path get() {
        init();
        // 非白名单图片格式返回null
        return canWrite() ? imagePath : null;
    }

    /**
     * 使用一段文本创建水印
     *
     * @param mark 文本
     * @return 水印对象
     */
    public static Watermark of(String mark) {
        return new Watermark(mark);
    }

    /**
     * 使用本地图片创建水印
     *
     * @param path 图片的本地文件
     * @return 水印对象Watermark
     */
    public static Watermark of(Path path) {
        return new Watermark(path);
    }

    /**
     * 使用图片流创建水印，可以下载远程图片创建水印
     *
     * @param is 图片流
     * @return 水印对象Watermark
     * @throws IOException 读取流异常
     */
    public static Watermark of(InputStream is) throws IOException {
        return new Watermark(is);
    }

    /**
     * 使用指定文本制作一张水印图片
     *
     * @param watermark 一段有意义的文本
     * @return the temp image path
     */
    private Path createWatermark(String watermark) {
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
            g2d.setColor(new Color(230, 230, 230));
            g2d.setStroke(new BasicStroke(1));
            g2d.setFont(new java.awt.Font("华文细黑", java.awt.Font.ITALIC, 24));
            g2d.rotate(Math.toRadians(-25));
            g2d.drawString(watermark, 100, 250);
            // Setting alpha
            g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
            g2d.dispose();
            ImageIO.write(bi, "png", temp.toFile());
            return temp;
        } catch (IOException e) {
            // Ignore
            LOGGER.error("Create watermark error.", e);
        }
        return null;
    }

    /**
     * 创建临时文件
     *
     * @return 临时文件路径
     * @throws IOException 没有权限或者磁盘不足等情况
     */
    private Path createTemp() throws IOException {
        temp = true;
        return Files.createTempFile("eec+watermark", "png");
    }

    /**
     * 删除临时文件，传入InputStream或文本时会保存到临时文件，所以需要清理资源
     *
     * @return 出异常时返回false
     */
    public boolean delete() {
        if (imagePath != null && temp) FileUtil.rm(imagePath);
        return true;
    }

    /**
     * 获取水印图片的后缀
     *
     * @return 水印图片后缀
     */
    public String getSuffix() {
        init();
        return "." + signature.extension;
    }

    /**
     * 水印图片Content-type
     *
     * @return Content-type
     */
    public String getContentType() {
        init();
        return signature.contentType;
    }

    /**
     * 测试水印图片是否可输出
     *
     * @return true: 资源可信任输出到Excel
     */
    public boolean canWrite() {
        init();
        return signature.isTrusted();
    }

    private void init() {
        if (signature == null) {
            if (imagePath == null && StringUtil.isNotEmpty(txt)) imagePath = createWatermark(txt);
            if (imagePath != null && signature == null) signature = FileSignatures.test(imagePath);
            if (signature == null) signature = new FileSignatures.Signature(null, "image/unknown");
        }
    }
}
