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

import org.junit.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;


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

        new Workbook("Picture test").addSheet(new ListSheet<>(list).ignoreHeader()
            .setColumns(new Column().setClazz(Path.class).setWidth(20)).setRowHeight(100))
            .writeTo(defaultTestPath);
    }

}
