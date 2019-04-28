/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.entity;

import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Created by guanquan.wang at 2018-01-29 16:05
 */
public class EmptySheet extends Sheet {
    public EmptySheet(Workbook workbook, String name, Column ... columns) {
        super(workbook, name, columns);
    }

    public EmptySheet(Workbook workbook, String name, WaterMark waterMark, Column ... columns) {
        super(workbook, name, waterMark, columns);
    }


//    @Override
//    public void writeTo(Path xl) throws IOException {
//        Path worksheets = xl.resolve("worksheets");
//        if (!Files.exists(worksheets)) {
//            FileUtil.mkdir(worksheets);
//        }
//        String name = getFileName();
//        workbook.what("0010", getName());
//
//
//        File sheetFile = worksheets.resolve(name).toFile();
//
////        // write date
////        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
////            // Write header
////            writeBefore(bw);
////            // Main data
////            // write ten empty rows
////            for (int i = 0; i < 10; i++) {
////                writeEmptyRow(bw);
////            }
////
////            // Write foot
////            writeAfter(bw);
////
////        } finally {
////            close();
////        }
////
////        // resize columns
////        boolean resize = false;
////        for  (Column hc : columns) {
////            if (hc.getWidth() > 0.000001) {
////                resize = true;
////                break;
////            }
////        }
////        if (getAutoSize() == 1 || resize) {
////            autoColumnSize(sheetFile);
////        }
////
////        // relationship
////        relManager.write(worksheets, name);
//    }

    @Override
    public RowBlock nextBlock() {
        return null;
    }

    /**
     * Returns total rows in this worksheet
     * @return 0
     */
    public int size() {
        return 0;
    }
}
