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

import org.dom4j.Document;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.io.Writer;
import java.nio.channels.Channel;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;

/**
 * File operation util.
 * Ignore the {@link IOException} and output error logs
 *
 * @author guanquan.wang on 2017-9-10
 */

public class FileUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(FileUtil.class);

    private FileUtil() { }


    /**
     * Close the {@link InputStream}
     *
     * @param inputStream the in stream
     */
    public static void close(InputStream inputStream) {
        if (inputStream != null) try {
            inputStream.close();
        } catch (IOException e) {
            LOGGER.error("close InputStream fail.", e);
        }
    }

    /**
     * Close the {@link OutputStream}
     *
     * @param outputStream the out stream
     */
    public static void close(OutputStream outputStream) {
        if (outputStream != null) {
            try {
                outputStream.close();
            } catch (IOException e) {
                LOGGER.error("close OutputStream fail.", e);
            }
        }
    }

    /**
     * Close the {@link Reader}
     *
     * @param br the reader
     */
    public static void close(Reader br) {
        if (br != null) {
            try {
                br.close();
            } catch (IOException e) {
                LOGGER.error("close Reader fail.", e);
            }
        }
    }

    /**
     * Close the {@link Writer}
     *
     * @param bw the writer
     */
    public static void close(Writer bw) {
        if (bw != null) {
            try {
                bw.close();
            } catch (IOException e) {
                LOGGER.error("close Writer fail.", e);
            }
        }
    }

    /**
     * Close the {@link Channel}
     *
     * @param channel the channel
     */
    public static void close(Channel channel) {
        if (channel != null) {
            try {
                channel.close();
            } catch (IOException e) {
                LOGGER.error("close Channel fail.", e);
            }
        }
    }

    /**
     * Delete if exists file
     *
     * @param path the file path to be delete
     */
    public static void rm(Path path) {
        rm(path.toFile());
    }

    /**
     * Delete if exists file
     *
     * @param file the file to be delete
     */
    public static void rm(File file) {
        if (file.exists()) {
            boolean boo = file.delete();
            if (!boo) {
                LOGGER.error("Delete file [{}] fail.", file.getPath());
            }
        }
    }

    /**
     * Remove file and sub-files if it a directory
     *
     * @param root the root path
     */
    public static void rm_rf(Path root) {
        rm_rf(root.toFile(), true);
    }

    /**
     * Remove file and sub-files if it a directory
     *
     * @param root the root path
     * @param rmSelf remove self if true
     */
    public static void rm_rf(File root, boolean rmSelf) {
        File temp = root;
        if (root.isDirectory()) {
            File[] subFiles = root.listFiles();
            if (subFiles == null) return;
            List<File> files = new ArrayList<>();
            int index = 0;
            do {
                files.addAll(Arrays.asList(subFiles));
                for (; index < files.size(); index++) {
                    if (files.get(index).isDirectory()) {
                        root = files.get(index);
                        subFiles = root.listFiles();
                        if (subFiles != null) {
                            files.addAll(Arrays.asList(subFiles));
                        }
                    }
                }
            } while (index < files.size());

            for (; --index >= 0; ) {
                rm(files.get(index));
            }
        }
        if (rmSelf) {
            rm(temp);
        }
    }

    /**
     * Copy a single file
     *
     * @param srcPath  the source path
     * @param descFile the destination file
     */
    public static void cp(Path srcPath, File descFile) {
        cp(srcPath.toFile(), descFile);
    }

    /**
     * Copy a single file
     *
     * @param srcFile  the source file
     * @param descFile the destination file
     */
    public static void cp(File srcFile, File descFile) {
        if (srcFile.length() == 0L) {
            try {
                boolean boo = descFile.createNewFile();
                if (!boo)
                    LOGGER.error("Copy file from [{}] to [{}] failed...", srcFile.getPath(), descFile.getPath());
                return;
            } catch (IOException e) {
            }
        }
        try (FileChannel inChannel = new FileInputStream(srcFile).getChannel();
             FileChannel outChannel = new FileOutputStream(descFile).getChannel()) {

            inChannel.transferTo(0, inChannel.size(), outChannel);
        } catch (IOException e) {
            LOGGER.error("Copy file from [{}] to [{}] failed...", srcFile.getPath(), descFile.getPath());
        }
    }

    /**
     * Copy a single file
     *
     * @param is  the source input stream
     * @param descFile the destination path
     */
    public static void cp(InputStream is, Path descFile) {
        try {
            Files.copy(is, descFile, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            LOGGER.error("Copy file to [{}] failed...", descFile);
        }
    }

    /**
     * Copy a directory to a new path
     *
     * @param srcPath       the source path string
     * @param descPath      the destination path string
     * @param moveSubFolder need move sub-folders
     */
    public static void cp_R(String srcPath, String descPath, boolean moveSubFolder) {
        copyFolder(srcPath, descPath, moveSubFolder);
    }

    /**
     * Copy a directory to a new path
     *
     * @param srcPath       the source path
     * @param descPath      the destination path
     * @param moveSubFolder need move sub-folders
     */
    public static void cp_R(Path srcPath, Path descPath, boolean moveSubFolder) {
        copyFolder(srcPath.toString(), descPath.toString(), moveSubFolder);
    }

    /**
     * Copy a directory to a new path
     *
     * @param srcPath       the source path string
     * @param descPath      the destination path string
     */
    public static void cp_R(String srcPath, String descPath) {
        copyFolder(srcPath, descPath, true);
    }

    /**
     * Copy a directory to a new path
     *
     * @param srcPath       the source path
     * @param descPath      the destination path
     */
    public static void cp_R(Path srcPath, Path descPath) {
        copyFolder(srcPath.toString(), descPath.toString(), true);
    }

    /**
     * Copy a directory to a new path
     *
     * @param srcPath       the source path string
     * @param descPath      the destination path string
     * @param moveSubFolder need move sub-folders
     */
    public static void copyFolder(String srcPath, String descPath, boolean moveSubFolder) {
        File src = new File(srcPath), desc = new File(descPath);
        if (!src.exists() || !src.isDirectory()) {
            throw new RuntimeException("The source path [" + srcPath + "] not exists or not a directory.");
        }
        if (!desc.exists()) {
            boolean boo = desc.mkdirs();
            if (!boo) {
                throw new RuntimeException("Create destination path [" + descPath + "] failed.");
            }
        }

        String ss[] = src.list();
        if (ss == null) return;
        List<File> files = new ArrayList<>();
        LinkedList<File> folders = new LinkedList<>();
        for (String s : ss) {
            File f = new File(src, s);
            if (f.isFile()) files.add(f);
            else folders.push(f);
        }
        ss = null;
        int src_path_len = srcPath.length();
        // If the source folder is not empty and need to copy sub-folders
        // Scan all sub-folders
        if (moveSubFolder && !folders.isEmpty()) {
            // 1. Scan all sub-folders, do not take recursion here
            while (!folders.isEmpty()) {
                File f = folders.pollLast();
                if (f == null) continue;
                File df = new File(desc, f.getPath().substring(src_path_len));
                // 1.1 Scan all sub-folders and create destination folders
                if (!df.exists() && !df.mkdir()) {
                    LOGGER.warn("Create sub-folder [{}] error skip it.", df.getPath());
                    continue;
                }
                File[] fs = f.listFiles();
                if (fs == null) continue;
                // 1.2 Storage the files witch need to copy
                for (File _f : fs) {
                    if (_f.isFile()) files.add(_f);
                    else folders.push(_f);
                }
            }
        }
        LOGGER.debug("Finished Scan. There contains {} files. Ready to copy them...", files.size());
        // 2. Copy files
        files.parallelStream().forEach(f -> cp(f, new File(descPath + f.getPath().substring(src_path_len))));
        LOGGER.debug("Copy all files in path {} finished.", srcPath);
    }

    /**
     * Create a directory
     *
     * @param destPath the destination directory path
     * @return the temp directory path
     * @throws IOException if I/O error occur
     */
    public static Path mkdir(Path destPath) throws IOException {
        Path path;
        if (isWindows()) {
            path = Files.createDirectories(destPath);
        } else {
            path = Files.createDirectories(destPath
                , PosixFilePermissions.asFileAttribute(PosixFilePermissions.fromString("rwxr-x---")));
        }
        return path;
    }

    /**
     * Create a temp directory
     *
     * @param prefix the directory prefix
     * @return the temp directory path
     * @throws IOException if I/O error occur
     */
    public static Path mktmp(String prefix) throws IOException {
        Path path;
        if (isWindows()) {
            path = Files.createTempDirectory(prefix);
        } else {
            path = Files.createTempDirectory(prefix
                , PosixFilePermissions.asFileAttribute(PosixFilePermissions.fromString("rwxr-x---")));
        }
        return path;
    }

    /**
     * Test current OS system is windows family
     *
     * @return true if OS is windows family
     */
    public static boolean isWindows() {
        return System.getProperty("os.name").toUpperCase().startsWith("WINDOWS");
    }

    /**
     * Write the {@link org.dom4j.Document} to a specify {@link Path}
     * with xml format
     *
     * @param doc the {@link org.dom4j.Document}
     * @param path the output path
     * @throws IOException if I/O error occur
     */
    public static void writeToDisk(Document doc, Path path) throws IOException {
        if (!exists(path.getParent())) {
            Files.createDirectories(path.getParent());
        }
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            //write the created document to an arbitrary file

            OutputFormat format = OutputFormat.createPrettyPrint();
            XMLWriter writer = new ExtXMLWriter(fos, format);
            writer.write(doc);
            writer.flush();
            writer.close();
        }
    }

    /**
     * Write the {@link org.dom4j.Document} to a specify {@link Path}
     *
     * @param doc the {@link org.dom4j.Document}
     * @param path the output path
     * @throws IOException if I/O error occur
     */
    public static void writeToDiskNoFormat(Document doc, Path path) throws IOException {
        if (!exists(path.getParent())) {
            mkdir(path.getParent());
        }
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            //write the created document to an arbitrary file

            XMLWriter writer = new ExtXMLWriter(fos);
            writer.write(doc);
            writer.flush();
            writer.close();
        }
    }

    /**
     * Tests whether a file exists.
     *
     * @param path the path to the file to test
     * @return false if not exits or no read permission
     */
    public static boolean exists(Path path) {
        try {
            return Files.exists(path);
        } catch (SecurityException e) {
            return false;
        }
    }
}
