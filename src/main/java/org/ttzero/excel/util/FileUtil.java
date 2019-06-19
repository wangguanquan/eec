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

package org.ttzero.excel.util;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.dom4j.Document;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

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
 * 文件操作工具类
 * <p>
 * Created by guanquan.wang on 2017-9-10
 */

public class FileUtil {
    private static Logger logger = LogManager.getLogger(FileUtil.class);

    private FileUtil() { }


    /**
     * 关闭输入流
     *
     * @param inputStream 文件输入流
     */
    public static void close(InputStream inputStream) {
        if (inputStream != null) try {
            inputStream.close();
        } catch (IOException e) {
            logger.error("close InputStream fail.", e);
        }
    }

    /**
     * 关闭输出流
     *
     * @param outputStream 文件输出流
     */
    public static void close(OutputStream outputStream) {
        if (outputStream != null) {
            try {
                outputStream.close();
            } catch (IOException e) {
                logger.error("close OutputStream fail.", e);
            }
        }
    }

    /**
     * 关闭BufferedReader
     *
     * @param br BufferedReader
     */
    public static void close(Reader br) {
        if (br != null) {
            try {
                br.close();
            } catch (IOException e) {
                logger.error("close Reader fail.", e);
            }
        }
    }

    /**
     * 关闭BufferedWriter
     *
     * @param bw BufferedWriter
     */
    public static void close(Writer bw) {
        if (bw != null) {
            try {
                bw.close();
            } catch (IOException e) {
                logger.error("close Writer fail.", e);
            }
        }
    }

    /**
     * 关闭Channel
     *
     * @param channel 通道
     */
    public static void close(Channel channel) {
        if (channel != null) {
            try {
                channel.close();
            } catch (IOException e) {
                logger.error("close Channel fail.", e);
            }
        }
    }

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
                logger.error("Delete file [{}] fail.", file.getPath());
            }
        }
    }

    public static void rm_rf(Path root) {
        rm_rf(root.toFile(), true);
    }

    /**
     * Remove file and sub files
     *
     * @param root the root path
     * @param rmSelf Remove self if true
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

    public static void cp(Path srcPath, File descFile) {
        cp(srcPath.toFile(), descFile);
    }

    /**
     * 复制单个文件
     *
     * @param srcFile  源文件
     * @param descFile 目标文件
     */
    public static void cp(File srcFile, File descFile) {
        if (srcFile.length() == 0L) {
            try {
                boolean boo = descFile.createNewFile();
                if (!boo)
                    logger.error("Copy file from [{}] to [{}] failed...", srcFile.getPath(), descFile.getPath());
                return;
            } catch (IOException e) {
            }
        }
        FileChannel inChannel = null, outChannel = null;
        try (FileInputStream fis = new FileInputStream(srcFile);
             FileOutputStream fos = new FileOutputStream(descFile)) {
            inChannel = fis.getChannel();
            outChannel = fos.getChannel();

            inChannel.transferTo(0, inChannel.size(), outChannel);
        } catch (IOException e) {
            logger.error("Copy file from [{}] to [{}] failed...", srcFile.getPath(), descFile.getPath());
        } finally {
            if (inChannel != null) {
                try {
                    inChannel.close();
                } catch (IOException e) {
                }
            }
            if (outChannel != null) {
                try {
                    outChannel.close();
                } catch (IOException e) {
                }
            }
        }
    }

    public static void cp(InputStream is, Path descFile) {
        try {
            Files.copy(is, descFile, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            logger.error("Copy file to [{}] failed...", descFile);
        }
    }

    public static void cp_R(String srcPath, String descPath, boolean moveSubFolder) {
        copyFolder(srcPath, descPath, moveSubFolder);
    }

    public static void cp_R(Path srcPath, Path descPath, boolean moveSubFolder) {
        copyFolder(srcPath.toString(), descPath.toString(), moveSubFolder);
    }

    public static void cp_R(String srcPath, String descPath) {
        copyFolder(srcPath, descPath, true);
    }

    public static void cp_R(Path srcPath, Path descPath) {
        copyFolder(srcPath.toString(), descPath.toString(), true);
    }

    /**
     * 复制整个文件夹的内容
     *
     * @param srcPath       源目录
     * @param descPath      目标目录
     * @param moveSubFolder 是否需要移动子文件夹
     */
    public static void copyFolder(String srcPath, String descPath, boolean moveSubFolder) {
        File src = new File(srcPath), desc = new File(descPath);
        // 源文件夹不存在
        if (!src.exists() || !src.isDirectory()) {
            throw new RuntimeException("源目录[" + srcPath + "]不是存在或者不是一个文件夹.");
        }
        // 目标文件夹不存在
        if (!desc.exists()) {
            boolean boo = desc.mkdirs();
            if (!boo) {
                throw new RuntimeException("目标文件夹[" + descPath + "]无法创建.");
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
        // 如果需要复制子文件夹并且源文件夹有子文件夹
        if (moveSubFolder && !folders.isEmpty()) {
            // 1. 扫描所有子文件夹, 这里不采取递归
            while (!folders.isEmpty()) {
                File f = folders.pollLast(), df = new File(desc, f.getPath().substring(src_path_len));
                // 1.1 扫描的同时为目标文件夹创建目录
                if (!df.exists() && !df.mkdir()) {
                    logger.warn("创建子文件夹[{}]失败跳过.", df.getPath());
                    continue;
                }
                File[] fs = f.listFiles();
                if (fs == null) continue;
                // 1.2 将文件及目标目录保存
                for (File _f : fs) {
                    if (_f.isFile()) files.add(_f);
                    else folders.push(_f);
                }
            }
        }
        logger.debug("扫描完成. 共计[{}]个文件. 开始复制文件...", files.size());
        // 2. 复制文件
        files.parallelStream().forEach(f -> cp(f, new File(descPath + f.getPath().substring(src_path_len))));
        logger.debug("复制结束.");
    }

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

    public static boolean isWindows() {
        return System.getProperty("os.name").toUpperCase().startsWith("WINDOWS");
    }

    public static void writeToDisk(Document doc, Path path) throws IOException {
        if (!Files.exists(path.getParent())) {
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

    public static void writeToDiskNoFormat(Document doc, Path path) throws IOException {
        if (!Files.exists(path.getParent())) {
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
}