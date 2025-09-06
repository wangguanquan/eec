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

package org.ttzero.excel.entity;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.SAXReaderUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Properties;
import java.util.stream.Stream;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * 工作薄输出协议，负责协调所有部件输出并组装所有零散的文件
 *
 * @author guanquan.wang at 2019-04-22 16:00
 * @see org.ttzero.excel.entity.e7.XMLWorkbookWriter
 * @see org.ttzero.excel.entity.csv.CSVWorkbookWriter
 */
public interface IWorkbookWriter extends Storable, Closeable {

    /**
     * 设置工作薄
     *
     * @param workbook 工作薄
     */
    void setWorkbook(Workbook workbook);

    /**
     * 获取最终的输出格式
     *
     * @return xlsx: excel07, xls: excel03
     */
    String getSuffix();

    /**
     * 将工作薄写到指定流
     *
     * @param os 输出流
     * @throws IOException if I/O error occur
     */
    void writeTo(OutputStream os) throws IOException;

    /**
     * 将工作薄写到指定位置
     *
     * @param file 目标文件
     * @throws IOException if I/O error occur
     */
    default void writeTo(File file) throws IOException {
        writeTo(file.toPath());
    }

    /**
     * 获取工作表输出协议
     *
     * @param sheet 工作表
     * @return 工作表输出协议
     */
    IWorksheetWriter getWorksheetWriter(Sheet sheet);

    @Deprecated
    default Path reMarkPath(Path source, Path target, String defaultName) throws IOException {
        return moveToPath(source, target, defaultName);
    }

    /**
     * 移动文件到指定位置，如果已存在相同文件名则会在文件名后追回{@code （n）}以区分，
     * {@code n}从1开始如果已存在{@code （n）}则新文件名为{@code （n + 1）}
     *
     * <p>例：目标文件夹已存在{@code a.xlsx}和{@code b（5）.xlsx}两个文件，添加一个名为{@code a.xlsx}的文件
     * 因为{@code a.xlsx}已存在所以新文件另存为{@code a（1）.xlsx}，添加一个名为{@code b.xlsx}的文件则新文件另存为{@code b（6）.xlsx}</p>
     *
     * @param source      源文件
     * @param target      目标文件夹
     * @param defaultName 目标文件名
     * @return 另存为目标文件绝对路径
     * @throws IOException if I/O error occur
     */
    default Path moveToPath(Path source, Path target, String defaultName) throws IOException {
        final String suffix = getSuffix();
        final boolean writeToFile = target.getFileName().toString().indexOf('.') > 0;
        Path out;
        if (writeToFile) {
            String fileName = target.getFileName().toString();
            if (!fileName.endsWith(suffix)) {
                fileName = fileName + suffix;
                out = target.getParent().resolve(fileName);
            } else out = target;
        } else out = target.resolve(defaultName + suffix);

        // 只指定文件夹时需判断文件名是否已存在，存在时添加编号导出
        if (!writeToFile && exists(out)) {
            Path parent = out.getParent();
            if (parent != null && exists(parent)) {
                final String matchPrefix = defaultName + " (", matchSuffix = ")" + suffix;
                long len = 1L;
                try (Stream<Path> pathStream = Files.list(parent)) {
                    len += pathStream.filter(p -> p.toFile().isFile() && p.getFileName().toString().startsWith(matchPrefix) && p.getFileName().toString().endsWith(matchSuffix)).count();
                } catch (Exception ex) {
                    // Ignore
                }
                String new_name;
                do {
                    new_name = matchPrefix + len++ + matchSuffix;
                } while (Files.exists(parent.resolve(new_name)));
                out = parent.resolve(new_name);
            }
        }
        Path parent = out.getParent();
        if (parent != null && !Files.exists(parent)) FileUtil.mkdir(parent);
        // Rename to xlsx
        Files.move(source, out, StandardCopyOption.REPLACE_EXISTING);
        return out;
    }

    /**
     * 获取pom配置相关信息
     *
     * @return general pom properties
     */
    static Properties pom() {
        Properties pom = new Properties();
        try {
            InputStream is = IWorkbookWriter.class.getClassLoader()
                .getResourceAsStream("META-INF/maven/org.ttzero/eec/pom.properties");
            if (is == null) {
                // Read from target/maven-archiver/pom.properties
                URL url = IWorkbookWriter.class.getClassLoader().getResource(".");
                if (url != null) {
                    Path targetPath = (FileUtil.isWindows()
                        ? Paths.get(url.getFile().substring(1))
                        : Paths.get(url.getFile())).getParent();
                    // On Mac or Linux
                    Path pomPath = targetPath.resolve("maven-archiver/pom.properties");
                    if (exists(pomPath)) {
                        is = Files.newInputStream(pomPath);
                        // On windows
                    } else {
                        pomPath = targetPath.getParent().resolve("pom.xml");
                        // load workbook.xml
                        SAXReader reader = SAXReaderUtil.createDefault();
                        Document document;
                        try {
                            document = reader.read(Files.newInputStream(pomPath));
                            Element pomRoot = document.getRootElement();
                            pom.setProperty("groupId", pomRoot.elementText("groupId"));
                            pom.setProperty("artifactId", pomRoot.elementText("artifactId"));
                            pom.setProperty("version", pomRoot.elementText("version"));
                        } catch (DocumentException | IOException e) {
                            // Nothing
                        }
                    }
                }
            }
            if (is != null) {
                pom.load(is);
            }
        } catch (IOException e) {
            // Nothing
        }
        if (StringUtil.isEmpty(pom.getProperty("version"))) {
            pom.setProperty("groupId", "org.ttzero");
            pom.setProperty("artifactId", "eec");
            pom.setProperty("version", "1.0.0");
        }
        return pom;
    }
}
