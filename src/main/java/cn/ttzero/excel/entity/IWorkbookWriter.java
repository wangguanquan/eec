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

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;

/**
 * Create by guanquan.wang at 2019-04-22 16:00
 */
public interface IWorkbookWriter {


    /**
     * The Workbook suffix
     * @return xlsx if excel07, xls if excel03
     */
    String getSuffix();

    /**
     * Write the workbook file to ${path}
     * @param path the storage path
     * @throws IOException if io error occur
     * @throws ExcelWriteException if error occur when excel write to file
     */
    void writeTo(Path path) throws IOException, ExcelWriteException;

    /**
     * Write to OutputStream ${os}
     * @param os the out put stream
     * @throws IOException if io error occur
     * @throws ExcelWriteException if error occur when excel write to file
     */
    void writeTo(OutputStream os) throws IOException, ExcelWriteException;

    /**
     * Write to file ${file}
     * @param file the storage file
     * @throws IOException if io error occur
     * @throws ExcelWriteException if error occur when excel write to file
     */
    void writeTo(File file) throws IOException, ExcelWriteException;

    /**
     * Write with template
     * @throws IOException if io error occur
     */
    Path template() throws IOException;
}
