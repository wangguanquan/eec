/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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

import okhttp3.Call;
import okhttp3.Callback;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;
import okhttp3.ResponseBody;
import okhttp3.ConnectionPool;
import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.ExcelColumn.ColType;
import org.ttzero.excel.drawing.PresetPictureEffect;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.util.FileSignatures;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.concurrent.TimeUnit;

import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2023-03-20 21:12
 */
public class PictureTest extends WorkbookTest {
    @Test public void testExportPicture() throws IOException {
        new Workbook("Picture test")
            .addSheet(new ListSheet<>(getLocalImages()).setColumns(new Column().setClazz(Path.class).writeAsMedia().setWidth(20)).setRowHeight(100))
            .writeTo(defaultTestPath);
    }

    @Test public void testBase64Image() throws IOException {
        new Workbook("Base64 image").addSheet(new ListSheet<>(Collections.singletonList("data:image/gif;base64,R0lGODlhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs="))
            .setColumns(new Column().setClazz(String.class).setWidth(20).writeAsMedia()).setRowHeight(100))
            .writeTo(defaultTestPath);
    }

    @Test public void testSyncRemoteImage() throws IOException {
        new Workbook("Sync download remote image").addSheet(new ListSheet<>(getRemoteUrls())
            .setColumns(new Column().setClazz(String.class).setWidth(20).writeAsMedia()).setRowHeight(100))
            .writeTo(defaultTestPath);
    }

    @Test public void testAsyncRemoteImage() throws IOException {
        new Workbook("Async download remote image").addSheet(new ListSheet<>(getRemoteUrls())
            .setColumns(new Column().setClazz(String.class).setWidth(20).writeAsMedia()).setRowHeight(100)
            .setSheetWriter(new XMLWorksheetWriter() {
            @Override
            public void downloadRemoteResource(Picture picture, String uri) {
                OkHttpClientUtil.client().newCall(new Request.Builder().url(uri).get().build()).enqueue(new Callback() {
                    @Override
                    public void onFailure(Call call, IOException e) {
                        try {
                            downloadCompleted(picture, null);
                        } catch (IOException ioException) {
                            ioException.printStackTrace();
                        }
                    }

                    @Override
                    public void onResponse(Call call, Response response) throws IOException {
                        try {
                            ResponseBody body;
                            if (response.isSuccessful() && (body = response.body()) != null) {
                                downloadCompleted(picture, body.bytes());
                            }
                        } catch (IOException ex) {
                            downloadCompleted(picture, null);
                        } finally {
                            response.close();
                        }
                    }
                });
            }
        })).writeTo(defaultTestPath);
    }

    @Test public void testStream() throws IOException {
        List<InputStream> list = getLocalImages().stream().map(p -> {
            try {
                return Files.newInputStream(p);
            } catch (IOException e) {
                return null;
            }
        }).filter(Objects::nonNull).collect(Collectors.toList());

        new Workbook("input-stream image").addSheet(new ListSheet<>(list)
            .setColumns(new Column().setClazz(InputStream.class).setWidth(20).writeAsMedia()).setRowHeight(100))
            .writeTo(defaultTestPath);
    }

    @Test public void testExportPictureAnnotation() throws IOException {
        new Workbook("test Picture annotation")
            .addSheet(new ListSheet<>(Pic.randomTestData()).setRowHeight(100))
            .writeTo(defaultTestPath);
    }

    @Test public void testPresetPictureEffects() throws IOException {
        new Workbook("Preset Picture Effects")
            .addSheet(new ListSheet<>(Pic2.randomTestData()).setRowHeight(217.5).setSheetWriter(new XMLWorksheetWriter() {
               private final Map<String, String> picCache = new HashMap<>();

                @Override
                protected Picture createPicture(int column, int row) {
                    Picture picture = super.createPicture(column, row);
                    picture.padding = 15 << 24 | 15 << 16 | 35 << 8 | 15;
                    if (row > 2) {
                        PresetPictureEffect[] effects = PresetPictureEffect.values();
                        effects[row - 3].preset(picture);
                    }
                    return picture;
                }

                @Override
                protected void writeFile(Path path, int row, int column, int xf) throws IOException {
                    writeNull(row, column, xf);
                    // Caching duplicate paths
                    String picName = picCache.get(path.toString());
                    if (picName != null) {
                        Picture picture = createPicture(column, row);
                        picture.picName = picName;
                        // Drawing
                        drawingsWriter.drawing(picture);
                        return;
                    }
                    // Test file signatures
                    FileSignatures.Signature signature = FileSignatures.test(path);
                    if ("unknown".equals(signature.extension)) {
                        LOGGER.warn("File types that are not allowed");
                        return;
                    }
                    int id = sheet.getWorkbook().incrementMediaCounter();
                    picName = "image" + id + "." + signature.extension;
                    // Store
                    Files.copy(path, mediaPath.resolve(picName), StandardCopyOption.REPLACE_EXISTING);

                    // Write picture
                    writePictureDirect(id, picName, column, row, signature);
                    picCache.put(path.toString(), picName);
                }

                final String[] name_cn_ZH = {"简单框架，白色","棱台亚光，白色","金属框架","矩形投影"
                    ,"映像圆角矩形","柔化边缘矩形","双框架，黑色","厚重亚光，黑色","简单框架，黑色","棱台形椭圆，黑色"
                    ,"复杂框架，黑色","中等复杂框架，黑色","居中矩形阴影","圆形对角，白色","剪去对角，白色"
                    ,"中等复杂框架，白色","旋转，白色","透视阴影，白色","松散透视，白色","柔化边缘椭圆","棱台矩形"
                    ,"棱台透视","映像右透视","棱台左透视，白色","映像棱台，黑色","映像棱台，白色","金属圆角矩形","金属椭圆",""};
                @Override
                protected void writeString(String s, int row, int column, int xf) throws IOException {
                    String ss;
                    try {
                        PresetPictureEffect effect = PresetPictureEffect.valueOf(s);
                        ss = name_cn_ZH[effect.ordinal()];
                    } catch (IllegalArgumentException ex) {
                        if ("None".equalsIgnoreCase(s)) ss = "无";
                        else ss = s;
                    }
                    super.writeString(ss, row, column, xf);
                }
            }))
            .writeTo(defaultTestPath);
    }

    public static class OkHttpClientUtil {

        private static class Handler {
            public static final OkHttpClient okHttpClient = new OkHttpClient.Builder()
                .retryOnConnectionFailure(true)
                .connectTimeout(60, TimeUnit.SECONDS)
                .readTimeout(60, TimeUnit.SECONDS)
                .writeTimeout(60, TimeUnit.SECONDS)
                .connectionPool(new ConnectionPool(20, 5L, TimeUnit.MINUTES))
                .hostnameVerifier((s, sslSession) -> true)
                .build();
        }

        OkHttpClientUtil() {
            Handler.okHttpClient.dispatcher().setMaxRequests(10);
            Handler.okHttpClient.dispatcher().setMaxRequestsPerHost(10);
        }

        public static OkHttpClient client() {
            return Handler.okHttpClient;
        }
    }

    static List<Path> getLocalImages() throws IOException {
        Path picturesPath = Paths.get(System.getProperty("user.home"), "Pictures");
        return Files.list(picturesPath).filter(p -> {
            String name = p.getFileName().toString();
            return !Files.isDirectory(p) && (name.endsWith(".png")
                || name.endsWith(".jpg") || name.endsWith(".webp")
                || name.endsWith(".wmf") || name.endsWith(".tif")
                || name.endsWith(".tiff") || name.endsWith(".gif")
                || name.endsWith(".jpeg") || name.endsWith(".ico")
                || name.endsWith(".emf") || name.endsWith(".bmp")
            );
        }).collect(Collectors.toList());
    }

    static List<String> getRemoteUrls() {
        return Arrays.asList("https://m.360buyimg.com/babel/jfs/t20260628/103372/21/40858/120636/649d00b3Fea336b50/1e97a70d3a3fe1c6.jpg"
            , "https://gw.alicdn.com/bao/uploaded/i3/1081542738/O1CN01ZBcPlR1W63BQXG5yO_!!0-item_pic.jpg_300x300q90.jpg"
            , "https://gw.alicdn.com/bao/uploaded/i3/2200754440203/O1CN01k8sRgC1DN1GGtuNT9_!!0-item_pic.jpg_300x300q90.jpg");
    }

    public static class Pic {
        @ExcelColumn("地址")
        private String addr;
        @ExcelColumn(value = "现场照片", colType = ColType.MEDIA)
        private String pic;

        public static List<Pic> randomTestData() {
            return getRemoteUrls().stream().map(u -> {
                Pic p = new Pic();
                p.addr = getRandomString();
                p.pic = u;
                return p;
            }).collect(Collectors.toList());
        }
    }

    public static class Pic2 {
        @ExcelColumn("Effect")
        private String effect;
        @ExcelColumn(value = "效果展示", colType = ColType.MEDIA, maxWidth = 53.75)
        private Path pic;

        public static List<Pic2> randomTestData() {
            Path path = testResourceRoot().resolve("elven-eyes.jpg");
            PresetPictureEffect[] effects = PresetPictureEffect.values();
            List<Pic2> list = new ArrayList<>(effects.length + 1);
            Pic2 pic = new Pic2();
            pic.effect = "None";
            pic.pic = path;
            list.add(pic);
            for (int i = 0; i < effects.length; i++) {
                pic = new Pic2();
                pic.effect = effects[i].name();
                pic.pic = path;
                list.add(pic);
            }
            return list;
        }
    }
}
