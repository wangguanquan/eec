/*
 * Copyright (c) 2017-2019, guanquan.wang@hotmail.com All Rights Reserved.
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

package org.ttzero.excel.entity.csv;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.ttzero.excel.entity.ExcelWriteException;
import org.ttzero.excel.entity.IPushModelSheet;
import org.ttzero.excel.entity.IWorkbookWriter;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;
import org.ttzero.excel.util.ZipUtil;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Write data as Comma-Separated Values format
 *
 * @author guanquan.wang at 2019-08-21 21:46
 */
public class CSVWorkbookWriter implements IWorkbookWriter {
    /**
     * LOGGER
     */
    protected Logger LOGGER = LoggerFactory.getLogger(getClass());
    protected Workbook workbook;
    /**
     * 单个文件时后缀为{@code csv}，多个文件时将多个csv归档为{@code zip}包
     */
    private String suffix = Const.Suffix.CSV;
    /**
     * Write BOM header
     */
    protected boolean withBom;
    /**
     * Charset 默认UTF-8
     */
    protected Charset charset;
    /**
     * 临时文件路径
     */
    protected Path tmpPath;

    public CSVWorkbookWriter() { }

    public CSVWorkbookWriter(Workbook workbook) {
        this.workbook = workbook;
    }

    public CSVWorkbookWriter(Workbook workbook, boolean withBom) {
        this.workbook = workbook;
        this.withBom = withBom;
    }

    @Override
    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * The Comma-Separated Values format suffix
     *
     * @return const value '.csv'
     */
    @Override
    public String getSuffix() {
        return suffix;
    }

    /**
     * Write to OutputStream
     *
     * @param os the output stream
     * @throws IOException         if io error occur
     */
    @Override
    public void writeTo(OutputStream os) throws IOException {
        Path path = null;
        try {
            path = createTemp();
            Files.copy(path, os);
        } finally {
            if (path != null) FileUtil.rm(path);
            close();
        }
    }

    @Override
    public void writeTo(Path root) throws IOException {
        Path path = null;
        try {
            path = createTemp();
            moveToPath(path, root);
        } finally {
            if (path != null) FileUtil.rm(path);
            close();
        }
    }

    protected void moveToPath(Path source, Path target) throws IOException {
        String name = StringUtil.isEmpty(workbook.getName()) ? "新建文件" : workbook.getName();
        Path resultPath = moveToPath(source, target, name);
        LOGGER.debug("Write completed. {}", resultPath);
    }

    @Override
    public Path writeBefore() throws IOException {
        if (tmpPath == null) {
            tmpPath = FileUtil.mktmp(Const.EEC_PREFIX);
            LOGGER.debug("Create temporary folder {}", tmpPath);
        }
        return tmpPath;
    }

    // Create csv file
    protected Path createTemp() throws IOException, ExcelWriteException {
        Path root = writeBefore();
        // Write worksheet data one by one
        for (int i = 0; i < workbook.getSize(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            try {
                if (!IPushModelSheet.class.isAssignableFrom(sheet.getClass()) || sheet.size() <= 0) {
                    // Collect properties
                    sheet.forWrite();
                    // Write to desk
                    sheet.writeTo(root);
                }
            } finally {
                sheet.close();
            }
        }

        // Zip compress if multi worksheet occur
        if (workbook.getSize() > 1) {
            suffix = Const.Suffix.ZIP;
            Path zipFile = ZipUtil.zipExcludeRoot(root, workbook.getCompressionLevel(), root);
            LOGGER.debug("Compression completed. {}", zipFile);
            return zipFile;
        } else {
            return root.resolve(workbook.getSheetAt(0).getName() + Const.Suffix.CSV);
        }
    }

    @Override
    public void close() {
        if (tmpPath != null) FileUtil.rm_rf(tmpPath);
    }

    /**
     * 设置字符集
     *
     * @param charset Charset
     * @return 当前Writer
     */
    public CSVWorkbookWriter setCharset(Charset charset) {
        this.charset = charset;
        return this;
    }

    // --- Customize worksheet writer

    public IWorksheetWriter getWorksheetWriter(Sheet sheet) {
        return new CSVWorksheetWriter(sheet, withBom).setCharset(charset);
    }
}
