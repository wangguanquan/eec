/*
 * Copyright (c) 2017-2019, guanquan.wang@yandex.com All Rights Reserved.
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
import org.ttzero.excel.entity.IWorkbookWriter;
import org.ttzero.excel.entity.IWorksheetWriter;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;
import org.ttzero.excel.util.ZipUtil;

import java.io.File;
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
    // The csv suffix
    private String suffix = Const.Suffix.CSV;
    // Write BOM header
    protected boolean withBom;
    // Charset 默认UTF-8
    protected Charset charset;

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
     * @param os the out put stream
     * @throws IOException         if io error occur
     */
    @Override
    public void writeTo(OutputStream os) throws IOException {
        Path path = createTemp();
        Files.copy(path, os);
        cleanTmp(path);
    }

    /**
     * Write to file
     *
     * @param file the storage file
     * @throws IOException         if io error occur
     */
    @Override
    public void writeTo(File file) throws IOException {
        Path path = createTemp();
        FileUtil.cp(path, file);
        cleanTmp(path);
    }

    /**
     * The Comma-Separated Values format do not support.
     *
     * @return the template path
     * @throws UnsupportedOperationException not support
     */
    @Override
    public Path template() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void writeTo(Path root) throws IOException {
        Path path = createTemp();
        reMarkPath(path, root);
        cleanTmp(path);
    }

    /**
     * Clean tmp folders
     *
     * @param path the tmp path
     */
    private void cleanTmp(Path path) {
        int i = 0;
        for (Path sub : path) {
            i++;
            if (sub.toString().startsWith(Const.EEC_PREFIX)) {
                break;
            }
        }
        if (i < path.getNameCount()) {
            FileUtil.rm_rf(path.getRoot().resolve(path.subpath(0, i)).toFile(), true);
            LOGGER.debug("Clean up temporary files");
        }
    }

    protected void reMarkPath(Path src, Path path) throws IOException {
        String name;
        if (StringUtil.isEmpty(name = workbook.getName())) {
            name = "新建文件";
        }

        Path resultPath = reMarkPath(src, path, name);
        LOGGER.debug("Write completed. {}", resultPath);
    }

    // Create csv file
    protected Path createTemp() throws IOException, ExcelWriteException {
        Sheet[] sheets = workbook.getSheets();
        for (int i = 0; i < sheets.length; i++) {
            Sheet sheet = sheets[i];
            IWorksheetWriter worksheetWriter = getWorksheetWriter(sheet);
            sheet.setSheetWriter(worksheetWriter);
            sheet.setId(i + 1);
            // default worksheet name
            if (StringUtil.isEmpty(sheet.getName())) {
                sheet.setName("Sheet" + (i + 1));
            }
            // Set cell value and style processor
            sheet.setCellValueAndStyle(new CSVCellValueAndStyle());

            // Force export all fields
            if (workbook.getForceExport() > sheet.getForceExport() && ListSheet.class.isAssignableFrom(sheet.getClass())) {
                ((ListSheet<?>) sheet).forceExport();
            }
        }
        LOGGER.debug("Sheet initialization completed.");

        Path root = null;
        try {
            root = FileUtil.mktmp(Const.EEC_PREFIX);
            LOGGER.debug("Create temporary folder {}", root);

            // Write worksheet data one by one
            for (int i = 0; i < workbook.getSize(); i++) {
                Sheet e = workbook.getSheetAt(i);
                e.writeTo(root);
                e.close();
            }

            // Zip compress if multi worksheet occur
            if (workbook.getSize() > 1) {
                suffix = Const.Suffix.ZIP;
                Path zipFile = ZipUtil.zipExcludeRoot(root, root);
                LOGGER.debug("Compression completed. {}", zipFile);
                FileUtil.rm_rf(root.toFile(), true);
                return zipFile;
            } else {
                return root.resolve(workbook.getSheetAt(0).getName() + Const.Suffix.CSV);
            }
        } catch (IOException | ExcelWriteException e) {
            // remove temp path
            if (root != null) FileUtil.rm_rf(root);
            throw e;
        }
    }

    @Override
    public void close() throws IOException { }

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
