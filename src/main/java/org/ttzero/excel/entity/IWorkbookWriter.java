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

package org.ttzero.excel.entity;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

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

import static org.ttzero.excel.util.FileUtil.exists;
import static org.ttzero.excel.util.StringUtil.indexOf;

/**
 * @author guanquan.wang at 2019-04-22 16:00
 */
public interface IWorkbookWriter extends Storable {

    /**
     * Setting workbook
     *
     * @param workbook the global workbook context
     */
    void setWorkbook(Workbook workbook);

    /**
     * The Workbook suffix
     *
     * @return xlsx if excel07, xls if excel03
     */
    String getSuffix();

    /**
     * Write to OutputStream ${os}
     *
     * @param os the out put stream
     * @throws IOException         if io error occur
     */
    void writeTo(OutputStream os) throws IOException;

    /**
     * Write to file ${file}
     *
     * @param file the storage file
     * @throws IOException         if io error occur
     */
    void writeTo(File file) throws IOException;

    /**
     * Write with template
     *
     * @return the template path
     * @throws IOException if io error occur
     */
    Path template() throws IOException;

    /**
     * Move src file into output path
     *
     * @param src the src file
     * @param rootPath the output root path
     * @param fileName the output file name
     * @return the output file path
     * @throws IOException if I/O error occur
     */
    default Path reMarkPath(Path src, Path rootPath, String fileName) throws IOException {
        // If the file exists, add the subscript after the file name.
        String suffix = getSuffix();
        Path o = rootPath.resolve(fileName + suffix);
        if (exists(o)) {
            final String fname = fileName;
            Path parent = o.getParent();
            if (parent != null && exists(parent)) {
                String[] os = parent.toFile().list((dir, name) ->
                    new File(dir, name).isFile()
                        && name.startsWith(fname)
                        && name.endsWith(suffix)
                );
                String new_name;
                if (os != null) {
                    int len = os.length, n;
                    do {
                        new_name = fname + " (" + len++ + ")" + suffix;
                        n = indexOf(os, new_name);
                    } while (n > -1);
                } else {
                    new_name = fname + suffix;
                }
                o = parent.resolve(new_name);
            } else {
                // Rename to
                Files.move(src, o, StandardCopyOption.REPLACE_EXISTING);
                return o;
            }
        }
        // Rename to xlsx
        Files.move(src, o);
        return o;
    }

    /**
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
                        SAXReader reader = new SAXReader();
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
