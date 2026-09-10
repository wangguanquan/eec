/*
 * Copyright (c) 2017, guanquan.wang@hotmail.com All Rights Reserved.
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

import org.ttzero.excel.manager.Const;

import java.io.BufferedOutputStream;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * zip util
 *
 * @author guanquan.wang on 2017/10/13.
 */
public class ZipUtil {
    // ZIP64 uses the maximum 32-bit value as the size placeholder.
    private static final long ZIP64_MAGIC_VALUE = 0xFFFFFFFFL;
    // ZIP64 0x2D
    private static final byte ZIP64_VERSION = 45;

    /**
     * Compression level for middle compression.
     */
    public static final int MIDDLE_COMPRESSION = 5;

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
        return zip(destPath, true, MIDDLE_COMPRESSION, srcPath);
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
        return zipExcludeRoot(destPath, MIDDLE_COMPRESSION, srcPath);
    }

    /**
     * zip files exclude root path
     * command: zip destPath srcPath1 srcPath2 ...
     *
     * @param destPath the destination path
     * @param compressionLevel compression level
     * @param srcPath  the source path
     * @return the result zip file path
     * @throws IOException if error occur.
     */
    public static Path zipExcludeRoot(Path destPath, int compressionLevel, Path... srcPath) throws IOException {
        if (!destPath.toString().endsWith(Const.Suffix.ZIP)) {
            destPath = Paths.get(destPath + Const.Suffix.ZIP);
        }
        if (!exists(destPath.getParent())) {
            FileUtil.mkdir(destPath.getParent());
        }
        return zip(destPath, false, compressionLevel, srcPath);
    }

    /**
     * zip files include root path
     * command: zip destPath srcPath1 srcPath2 ...
     *
     * @param destPath     the destination path
     * @param compressRoot include root path if true
     * @param compressionLevel compression level
     * @param srcPath      the source path
     * @return the result zip file path
     * @throws IOException if error occur.
     */
    private static Path zip(Path destPath, boolean compressRoot, int compressionLevel, Path... srcPath) throws IOException {
        CountingOutputStream cos = new CountingOutputStream(
            new BufferedOutputStream(Files.newOutputStream(destPath, StandardOpenOption.CREATE)));
        ZipOutputStream zos = new ZipOutputStream(cos);
        zos.setLevel(Math.min(Math.max(compressionLevel, 0), 9));
        List<Path> paths = new ArrayList<>();
        // Local header offsets of entries whose uncompressed size requires ZIP64.
        List<Long> zip64Offsets = new ArrayList<>();
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
                // required Zip64.
                if (Files.size(paths.get(j)) >= ZIP64_MAGIC_VALUE) {
                    zip64Offsets.add(cos.getCount());
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
        // Patch after close so all buffered ZIP data has been written to disk.
        if (!zip64Offsets.isEmpty()) {
            patchZip64LocalHeaders(destPath, zip64Offsets);
        }
        return destPath;
    }

    // ZIP local file header signature: PK\003\004
    private static void patchZip64LocalHeaders(Path zipPath, List<Long> offsets) throws IOException {
        byte[] signature = new byte[4];
        try (SeekableByteChannel channel = Files.newByteChannel(zipPath, StandardOpenOption.READ, StandardOpenOption.WRITE)) {
            for (long offset : offsets) {
                ByteBuffer buffer = ByteBuffer.wrap(signature);
                channel.position(offset);
                // A channel read is not guaranteed to fill the buffer in one call.
                while (buffer.hasRemaining()) {
                    if (channel.read(buffer) < 0) {
                        break;
                    }
                }
                // Stop patching if an offset does not point to a complete local file header.
                if (buffer.hasRemaining() || signature[0] != 0x50 || signature[1] != 0x4B
                    || signature[2] != 0x03 || signature[3] != 0x04) {
                    break;
                }
                // version needed to extract is the two-byte field after the signature.
                channel.position(offset + 4);
                buffer = ByteBuffer.wrap(new byte[] { ZIP64_VERSION, 0x00 });
                while (buffer.hasRemaining()) channel.write(buffer);
            }
        }
    }

    // Tracks the exact local header offsets written by ZipOutputStream.
    private static final class CountingOutputStream extends FilterOutputStream {
        private long count;

        CountingOutputStream(OutputStream out) {
            super(out);
        }

        long getCount() {
            return count;
        }

        @Override
        public void write(int b) throws IOException {
            out.write(b);
            count++;
        }

        @Override
        public void write(byte[] b, int off, int len) throws IOException {
            out.write(b, off, len);
            count += len;
        }
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
            if (!exists(sub.getParent())) {
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
