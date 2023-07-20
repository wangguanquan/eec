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
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;
import java.util.concurrent.TimeUnit;

/**
 * @author wangguanquan3 at 2023-03-20 21:12
 */
public class PictureTest extends WorkbookTest {
    @Test public void testExportPicture() throws IOException {
        Path picturesPath = Paths.get(System.getProperty("user.home"), "Pictures");
        List<Path> list = Files.list(picturesPath).filter(p -> {
            String name = p.getFileName().toString();
            return !Files.isDirectory(p) && (name.endsWith(".png")
                || name.endsWith(".jpg") || name.endsWith(".webp")
                || name.endsWith(".wmf") || name.endsWith(".tif")
                || name.endsWith(".tiff") || name.endsWith(".gif")
                || name.endsWith(".jpeg") || name.endsWith(".ico")
                || name.endsWith(".emf") || name.endsWith(".bmp")
            );
        }).collect(Collectors.toList());

        new Workbook("Picture test")
            .addSheet(new ListSheet<>(list).setColumns(new Column().setClazz(Path.class).setWidth(20)).setRowHeight(100))
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

    static List<String> getRemoteUrls() {
        return Arrays.asList("https://m.360buyimg.com/babel/jfs/t20260628/103372/21/40858/120636/649d00b3Fea336b50/1e97a70d3a3fe1c6.jpg"
            , "https://gw.alicdn.com/bao/uploaded/i3/1081542738/O1CN01ZBcPlR1W63BQXG5yO_!!0-item_pic.jpg_300x300q90.jpg"
            , "https://gw.alicdn.com/bao/uploaded/i3/2200754440203/O1CN01k8sRgC1DN1GGtuNT9_!!0-item_pic.jpg_300x300q90.jpg");
    }
}
