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
import org.ttzero.excel.Print;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.MediaColumn;
import org.ttzero.excel.drawing.PresetPictureEffect;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.reader.Drawings;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.util.FileSignatures;

import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.concurrent.TimeUnit;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2023-03-20 21:12
 */
public class PictureTest extends WorkbookTest {
    @Test public void testExportPicture() throws IOException {
        List<Path> expectList = getLocalImages();
        new Workbook()
            .addSheet(new ListSheet<>(expectList).setColumns(new Column().writeAsMedia().setWidth(20)).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("Picture test (Path).xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Picture test (Path).xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Path expectPath = expectList.get(i);
                Drawings.Picture pic = list.get(i);
                // Check file size
                assertEquals(Files.size(expectPath), Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectPath), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testExportPictureUseFile() throws IOException {
        List<Path> expectList = getLocalImages();
        new Workbook()
            .addSheet(new ListSheet<>(expectList.stream().map(Path::toFile).collect(Collectors.toList())).setColumns(new Column().writeAsMedia().setWidth(20)).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("Picture test (File).xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Picture test (File).xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Path expectPath = expectList.get(i);
                Drawings.Picture pic = list.get(i);
                // Check file size
                assertEquals(Files.size(expectPath), Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectPath), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testExportPictureUseByteArray() throws IOException {
        List<Path> expectList = getLocalImages();
        new Workbook()
            .addSheet(new ListSheet<>(expectList.stream().map(e -> {
                try {
                    return Files.readAllBytes(e);
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
                return null;
            }).collect(Collectors.toList())).setColumns(new Column().writeAsMedia().setWidth(20)).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("Picture test (Byte array).xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Picture test (Byte array).xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Path expectPath = expectList.get(i);
                Drawings.Picture pic = list.get(i);
                // Check file size
                assertEquals(Files.size(expectPath), Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectPath), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testExportPictureUseBuffer() throws IOException {
        List<Path> expectList = getLocalImages();
        new Workbook()
            .addSheet(new ListSheet<>(expectList.stream().map(e -> {
                try (SeekableByteChannel channel = Files.newByteChannel(e, StandardOpenOption.READ)) {
                    ByteBuffer buffer = ByteBuffer.allocate((int) channel.size());
                    channel.read(buffer);
                    buffer.flip();
                    return buffer;
                } catch (IOException ex) {
                    return null;
                }
            }).collect(Collectors.toList())).setColumns(new Column().writeAsMedia().setWidth(20)).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("Picture test (Buffer).xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Picture test (Buffer).xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Path expectPath = expectList.get(i);
                Drawings.Picture pic = list.get(i);
                // Check file size
                assertEquals(Files.size(expectPath), Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectPath), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testExportPictureUseStream() throws IOException {
        List<Path> expectList = getLocalImages();
        List<InputStream> inputStreams = expectList.stream().map(p -> {
            try {
                return Files.newInputStream(p);
            } catch (IOException e) {
                return null;
            }
        }).filter(Objects::nonNull).collect(Collectors.toList());

        new Workbook().addSheet(new ListSheet<>(inputStreams)
            .setColumns(new Column().setWidth(20).writeAsMedia()).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("Picture test (InputStream).xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Picture test (InputStream).xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Path expectPath = expectList.get(i);
                Drawings.Picture pic = list.get(i);
                // Check file size
                assertEquals(Files.size(expectPath), Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectPath), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testBase64Image() throws IOException {
        String base64Image = "R0lGODlhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs=";
        new Workbook().addSheet(new ListSheet<>(Collections.singletonList("data:image/gif;base64," + base64Image))
            .setColumns(new Column().setWidth(20).writeAsMedia()).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("Base64 image.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Base64 image.xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(list.size(), 1);
            Drawings.Picture pic = list.get(0);
            // Check CRC32
            assertEquals(crc32(Base64.getDecoder().decode(base64Image)), crc32(pic.getLocalPath()));
        }
    }

    @Test public void testSyncRemoteImage() throws IOException {
        List<String> expectList = getRemoteUrls();
        new Workbook().addSheet(new ListSheet<>(expectList)
            .setColumns(new Column().setWidth(20).writeAsMedia()).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("Sync download remote image.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Sync download remote image.xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Drawings.Picture pic = list.get(i);
                byte[] expectBytes = getRemoteData(expectList.get(i));
                // Check file size
                assertEquals(expectBytes.length, Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectBytes), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testSyncRemoteImageUseOkHTTP() throws IOException {
        List<String> expectList = getRemoteUrls();
        new Workbook().addSheet(new ListSheet<>(expectList)
            .setColumns(new Column().setWidth(20).writeAsMedia()).setRowHeight(100)
            .setSheetWriter(new XMLWorksheetWriter() {
                @Override public void downloadRemoteResource(Picture picture, String uri) throws IOException {
                    if (uri.startsWith("http")) {
                        try (Response response = OkHttpClientUtil.client().newCall(new Request.Builder().url(uri).get().build()).execute()) {
                            ResponseBody body;
                            if (response.isSuccessful() && (body = response.body()) != null) {
                                downloadCompleted(picture, body.bytes());
                            }
                        } catch (IOException ex) {
                            downloadCompleted(picture, null);
                        }
                    }
                    else if (uri.startsWith("ftp")) {
                        Print.println("down load from ftp server");
                    }
                }
            })).writeTo(defaultTestPath.resolve("sync download remote image use OkHttp.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("sync download remote image use OkHttp.xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Drawings.Picture pic = list.get(i);
                byte[] expectBytes = getRemoteData(expectList.get(i));
                // Check file size
                assertEquals(expectBytes.length, Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectBytes), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testAsyncRemoteImage() throws IOException {
        List<String> expectList = getRemoteUrls();
        new Workbook().addSheet(new ListSheet<>(expectList)
            .setColumns(new Column().setWidth(20).writeAsMedia()).setRowHeight(100)
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
        })).writeTo(defaultTestPath.resolve("Async download remote image.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Async download remote image.xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Drawings.Picture pic = list.get(i);
                byte[] expectBytes = getRemoteData(expectList.get(i));
                // Check file size
                assertEquals(expectBytes.length, Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectBytes), crc32(pic.getLocalPath()));
            }
        }
    }

    @Test public void testExportPictureAnnotation() throws IOException {
        List<Pic> expectList = Pic.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("test Picture annotation.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test Picture annotation.xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Drawings.Picture pic = list.get(i);
                byte[] expectBytes = getRemoteData(expectList.get(i).pic);
                // Check file size
                assertEquals(expectBytes.length, Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectBytes), crc32(pic.getLocalPath()));
            }

            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).dataRows().iterator();
            for (Pic p : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(p.addr, row.getString(0));
            }
        }
    }

    @Test public void testExportPictureAutoSize() throws IOException {
        List<Pic> expectList = Pic.randomTestData();
        new Workbook()
            .setAutoSize(true)
            .addSheet(new ListSheet<>(expectList).setRowHeight(100))
            .writeTo(defaultTestPath.resolve("test Picture auto-size.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test Picture auto-size.xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Drawings.Picture pic = list.get(i);
                byte[] expectBytes = getRemoteData(expectList.get(i).pic);
                // Check file size
                assertEquals(expectBytes.length, Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectBytes), crc32(pic.getLocalPath()));
            }

            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).dataRows().iterator();
            for (Pic p : expectList) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(p.addr, row.getString(0));
            }
        }
    }

    @Test public void testPresetPictureEffects() throws IOException {
        List<Pic2> expectList = Pic2.randomTestData();
        new Workbook()
            .addSheet(new ListSheet<>(expectList).setRowHeight(217.5).autoSize().setSheetWriter(new XMLWorksheetWriter() {
               private final Map<String, String> picCache = new HashMap<>();

                @Override
                protected Picture createPicture(int column, int row) {
                    Picture picture = super.createPicture(column, row);
                    picture.setPaddingTop(15).setPaddingRight(-15).setPaddingBottom(-35).setPaddingLeft(15);
                    PresetPictureEffect[] effects = PresetPictureEffect.values();
                    picture.effect = effects[row - 2].getEffect();
                    return picture;
                }

                @Override
                protected void writeFile(Path path, int row, int column) throws IOException {
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
            }))
            .writeTo(defaultTestPath.resolve("Preset Picture Effects.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Preset Picture Effects.xlsx"))) {
            List<Drawings.Picture> list = reader.sheet(0).listPictures();
            assertEquals(expectList.size(), list != null ? list.size() : 0);
            for (int i = 0; i < expectList.size(); i++) {
                Path expectPath = expectList.get(i).pic;
                Drawings.Picture pic = list.get(i);
                // Check file size
                assertEquals(Files.size(expectPath), Files.size(pic.getLocalPath()));
                // Check CRC32
                assertEquals(crc32(expectPath), crc32(pic.getLocalPath()));
            }

            Iterator<org.ttzero.excel.reader.Row> iter = reader.sheet(0).dataRows().iterator();
            PresetPictureEffect[] effects = PresetPictureEffect.values();

            for (PresetPictureEffect p : effects) {
                assertTrue(iter.hasNext());
                org.ttzero.excel.reader.Row row = iter.next();
                assertEquals(p.name(), row.getString(0));
            }
        }
    }

    @Test public void testExportPictureAutoSizePaging() throws IOException {
        List<Path> expectList = new ArrayList<>(256);
        for (int i = 0; i < 5; i++, expectList.addAll(getLocalImages()));

        IWorksheetWriter worksheetWriter;
        new Workbook()
            .addSheet(new ListSheet<>(expectList).setColumns(new Column().writeAsMedia().setWidth(20)).setRowHeight(100)
                .setSheetWriter(worksheetWriter = new XMLWorksheetWriter() {
                    @Override
                    public int getRowLimit() {
                        return 16;
                    }
                }))
            .writeTo(defaultTestPath.resolve("test Picture auto-size paging.xlsx"));

        int count = expectList.size(), rowLimit = worksheetWriter.getRowLimit();
        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("test Picture auto-size paging.xlsx"))) {
            if (expectList.size() > 0) {
                assertEquals(reader.getSheetCount(), (count % rowLimit > 0 ? count / rowLimit + 1 : count / rowLimit));

                for (int i = 0, len = reader.getSheetCount(), a = 0; i < len; i++) {
                    List<Drawings.Picture> list = reader.sheet(i).listPictures();
                    if (i < len - 1) assertEquals(list.size(), rowLimit);
                    else assertEquals(expectList.size() - rowLimit * (len - 1), list.size());
                    for (int j = 0; j < list.size(); j++) {
                        Path expectPath = expectList.get(a++);
                        Drawings.Picture pic = list.get(j);
                        // Check file size
                        assertEquals(Files.size(expectPath), Files.size(pic.getLocalPath()));
                        // Check CRC32
                        assertEquals(crc32(expectPath), crc32(pic.getLocalPath()));
                    }
                }
            } else assertNull(reader.listPictures());
        }
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

    public static byte[] getRemoteData(String url) {
        try (Response response = OkHttpClientUtil.client().newCall(new Request.Builder().url(url).get().build()).execute()) {
            ResponseBody body;
            if (response.isSuccessful() && (body = response.body()) != null) {
                return body.bytes();
            }
        } catch (IOException ex) { }
        return new byte[] { };
    }

    static List<Path> getLocalImages() throws IOException {
        Path picturesPath = Paths.get(System.getProperty("user.home"), "Pictures");
        if (!Files.exists(picturesPath)) return Collections.emptyList();
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
        @MediaColumn(presetEffect = PresetPictureEffect.Rotated_White)
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
        @MediaColumn
        @ExcelColumn(value = "效果展示", maxWidth = 53.75)
        private Path pic;

        public static List<Pic2> randomTestData() {
            Path path = testResourceRoot().resolve("elven-eyes.jpg");
            PresetPictureEffect[] effects = PresetPictureEffect.values();
            List<Pic2> list = new ArrayList<>(effects.length);
            for (int i = 0; i < effects.length; i++) {
                Pic2 pic = new Pic2();
                pic.effect = effects[i].name();
                pic.pic = path;
                list.add(pic);
            }
            return list;
        }
    }
}
