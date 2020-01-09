/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.util;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.zip.CRC32;
import java.util.zip.CheckedOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * zip util
 *
 * @author guanquan.wang on 2017/10/13.
 */
public class ZipUtil {
    private static final String suffix = ".zip";

    private ZipUtil() { }

    /**
     * zip files exclude root path
     * command: zip destPath srcPath1 srcPath2 ...
     *
     * @param destPath the destination path
     * @param srcPath  the source path
     * @return the result zip file path
     * @throws IOException if error occur.
     */
    public static Path zip(Path destPath, Path... srcPath) throws IOException {
        return zip(destPath, true, srcPath);
    }

    /**
     * zip files exclude root path
     * command: zip destPath srcPath1 srcPath2 ...
     *
     * @param destPath the destination path
     * @param srcPath  the source path
     * @return the result zip file path
     * @throws IOException if error occur.
     */
    public static Path zipExcludeRoot(Path destPath, Path... srcPath) throws IOException {
        return zip(destPath, false, srcPath);
    }

    /**
     * zip files include root path
     * command: zip destPath srcPath1 srcPath2 ...
     *
     * @param destPath     the destination path
     * @param compressRoot include root path if true
     * @param srcPath      the source path
     * @return the result zip file path
     * @throws IOException if error occur.
     */
    private static Path zip(Path destPath, boolean compressRoot, Path... srcPath) throws IOException {
        if (!destPath.toString().endsWith(suffix)) {
            destPath = Paths.get(destPath.toString() + suffix);
        }
        if (!Files.exists(destPath.getParent())) {
            FileUtil.mkdir(destPath.getParent());
        }
        ZipOutputStream zos = new ZipOutputStream(new CheckedOutputStream(
            Files.newOutputStream(destPath, StandardOpenOption.CREATE), new CRC32()));
        List<Path> paths = new ArrayList<>();
        int i = 0, index = 0;
        int[] array = new int[srcPath.length];
        for (Path src : srcPath) {
            if (Files.isDirectory(src)) {
                paths.addAll(subPath(src));
                while (i < paths.size()) {
                    if (Files.isDirectory(paths.get(i))) {
                        paths.addAll(subPath(paths.get(i)));
                    }
                    i++;
                }
            } else {
                paths.add(src);
                i++;
            }
            array[index++] = i;
        }

        index = 0;
        Path basePath = compressRoot ? srcPath[index].getParent() : srcPath[index];
        for (int j = 0, len = basePath.toString().length(); j < i; j++) {
            if (Files.isDirectory(paths.get(j))) continue;
            if (j < array[index]) {
                String name;
                if (paths.get(j).equals(srcPath[index])) {
                    name = paths.get(j).getNameCount() > 1
                        ? paths.get(j).toString().substring(paths.get(j).getParent().toString().length() + 1)
                        : paths.get(j).toString();
                } else {
                    name = paths.get(j).toString().substring(len + 1);
                }
                zos.putNextEntry(new ZipEntry(name));
                Files.copy(paths.get(j), zos);
                zos.closeEntry();
            } else {
                basePath = compressRoot ? srcPath[++index].getParent() : srcPath[++index];
                len = basePath.toString().length();
                j--;
            }
        }

        zos.close();
        return destPath;
    }

    private static List<Path> subPath(Path path) throws IOException {
        try (Stream<Path> fileStream = Files.list(path)) {
            return fileStream.collect(Collectors.toList());
        }
    }

    /**
     * unzip file to descPath
     *
     * @param stream   the input stream
     * @param destPath the destination path
     * @return the result zip file path
     * @throws IOException if error occur.
     */
    public static Path unzip(InputStream stream, Path destPath) throws IOException {
        if (!Files.isDirectory(destPath)) {
            FileUtil.mkdir(destPath);
        }
        ZipInputStream zis = new ZipInputStream(stream);
        ZipEntry entry = zis.getNextEntry();
        while (entry != null) {
            Path sub = destPath.resolve(entry.getName());
            // Create parent
            if (!Files.exists(sub.getParent())) {
                FileUtil.mkdir(sub.getParent());
            }
            if (entry.isDirectory()) {
                FileUtil.mkdir(sub);
            } else {
                FileUtil.cp(zis, sub);
            }
            zis.closeEntry();
            entry = zis.getNextEntry();
        }

        zis.close();
        return destPath;
    }

}
