/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;


/**
 * FILE SIGNATURES
 *
 * @author guanquan.wang at 2023-02-13 16:26
 */
public class FileSignatures {
    /**
     * LOGGER
     */
    final static Logger LOGGER = LoggerFactory.getLogger(FileSignatures.class);
    /**
     * Configure trusted image types
     */
    public static Map<String, String> whitelist = new HashMap<String, String>() {{
        put("png", "image/png");
        put("jpg", "image/jpg");
        put("gif", "image/gif");
        put("tiff", "image/tiff");
        put("bmp", "image/bmp");
        put("ico", "image/x-ico");
        put("tif", "image/tiff");
        put("emf", "image/x-emf");
        put("wmf", "image/x-wmf");
        put("webp", "image/webp");
    }};
    private FileSignatures() { }

    public static Signature test(Path path) {
        Signature signature = null;
        try (InputStream is = Files.newInputStream(path)) {
            byte[] bytes = new byte[1 << 9];
            int n = is.read(bytes);
            signature = test(ByteBuffer.wrap(bytes, 0, n));
        } catch (Exception ex) {
            LOGGER.warn("Test file signature occur error.", ex);
        }
        if (signature == null) {
            String name = path.getFileName().toString();
            int i = name.lastIndexOf('.');
            String uncertainExtensionName = i > 0 && i < name.length() - 1 ? name.substring(i + 1) : "unknown";
            signature = new Signature(uncertainExtensionName, whitelist.getOrDefault(uncertainExtensionName, "image/unknown"), 0, 0);
        }
        return signature;
    }

    public static Signature test(ByteBuffer buffer) {
        if (buffer.remaining() < 32) return null;
        int t0 = buffer.getShort() & 0xFFFF;

        String extension = null;
        int width = 0, height = 0;

        // Maybe jpg
        if (t0 == 0xFFD8) {
            extension = "jpg";
            for (; buffer.remaining() >= 4; ) {
                int t1 = buffer.getShort() & 0xFFFF, n = buffer.getShort() & 0xFFFF;
                if (t1 == 0xFFC0) {
                    if (buffer.remaining() >= 5) {
                        buffer.get();
                        height = buffer.getShort() & 0xFFFF;
                        width  = buffer.getShort() & 0xFFFF;
                    }
                    break;
                }
                else if (buffer.remaining() >= n) buffer.position(buffer.position() + n - 2);
                else break;
            }
        }
        // BMP/DIB
        else if (t0 == 0x424D) {
            extension = "bmp";
            buffer.order(ByteOrder.LITTLE_ENDIAN);
            buffer.position(buffer.position() + 16);
            width  = buffer.getInt();
            height = buffer.getInt();
        }

        if (extension != null) {
            return new Signature(extension, whitelist.getOrDefault(extension, "image/unknown"), width, height);
        }

        buffer.position(0);
        t0 = buffer.getInt();

        byte v;
        switch (t0) {
            // Maybe PNG
            case 0x89504E47:
                int t1 = buffer.getInt();
                if (t1 == 0x0D0A1A0A) {
                    extension = "png";
                    buffer.getLong();
                    width  = buffer.getInt();
                    height = buffer.getInt();
                }
                break;
            // TIFF II.*
            case 0x49492A00: buffer.order(ByteOrder.LITTLE_ENDIAN);
            // TIFF MM.*
            case 0x4D4D002A:
            // TIFF MM.+ Tagged Image File Format files >4 GB
            case 0x4D4D002B:
                return tiff(buffer);
            // GIF
            case 0x47494638:
                extension = "gif";
                buffer.getShort();
                buffer.order(ByteOrder.LITTLE_ENDIAN);
                width  = buffer.getShort() & 0xFFFF;
                height = buffer.getShort() & 0xFFFF;
                break;
            // ICO
            case 0x100:
                extension = "ico";
                buffer.getShort();
                width  = (v = buffer.get()) != 0 ? v & 0xFF : 0x100;
                height = (v = buffer.get()) != 0 ? v & 0xFF : 0x100;
                break;
            // EMF
            case 0x01000000:
                extension = "emf";
                buffer.getInt(); // Ignore
                buffer.order(ByteOrder.LITTLE_ENDIAN);
                int left = buffer.getInt(), top = buffer.getInt(), right = buffer.getInt(), bottom = buffer.getInt();
                width  = Math.max(0, right - left + 1);
                height = Math.max(0, bottom - top + 1);
                break;
            // WMF
            case 0xD7CDC69A:
            case 0x01000900:
                extension = "wmf";
                buffer.order(ByteOrder.LITTLE_ENDIAN);
                buffer.getShort(); // Ignore
                left = buffer.getShort() & 0xFFFF;
                top = buffer.getShort() & 0xFFFF;
                right = buffer.getShort() & 0xFFFF;
                bottom = buffer.getShort() & 0xFFFF;
                int inch = buffer.getShort() & 0XFFFF;

                double coeff = inch > 0 ? 72.0D / inch : 1.0D;
                width  = (int) Math.round((right - left) * coeff);
                height = (int) Math.round((bottom - top) * coeff);
                break;
            // WEBP
            case 0x52494646:
                extension = "webp";
                buffer.order(ByteOrder.LITTLE_ENDIAN);
                // Chunk Size
                int size = buffer.getInt()
                    , x = 0x50424557; // ascii: webp
                if (buffer.getInt() == x) {
                    int chunkType = buffer.getInt(), blockSize = buffer.getInt();
                    switch (chunkType) {
                        // VP8
                        case 0x20385056:
                            int tmp = buffer.get() & 0xFF | (buffer.get() & 0xFF) << 8 | (buffer.get() & 0xFF) << 16;
//                            key_frame = tmp & 0x1;
//                            version = (tmp >> 1) & 0x7;
//                            show_frame = (tmp >> 4) & 0x1;
//                            first_part_size = (tmp >> 5) & 0x7FFFF;
                            // show_frame flag (0 when current frame is not for display,1 when current frame is for display).
//                            if ((tmp & 1) == 1) {
//                                byte version = (byte) ((tmp >> 1) & 0x7);
                                // Ignore others...
                            tmp = buffer.get() & 0xFF | (buffer.get() & 0xFF) << 8 | (buffer.get() & 0xFF) << 16;
                            if (tmp == 0x2a019d) {
                                width  = buffer.getShort() & 0x3FFF;
                                height = buffer.getShort() & 0x3FFF;
                            }
//                            }
                            break;
                        // VP8X
                        case 0x58385056:
                            buffer.getInt(); // Ignore
                            width  = 1 + (buffer.get() & 0xFF | (buffer.get() & 0xFF) << 8 | (buffer.get() & 0xFF) << 16);
                            height = 1 + (buffer.get() & 0xFF | (buffer.get() & 0xFF) << 8 | (buffer.get() & 0xFF) << 16);
                            break;
                    }

                }
                break;
            default:
        }

        if (extension != null) {
            return new Signature(extension, whitelist.getOrDefault(extension, "image/unknown"), width, height);
        }

        return null;
    }

    public static Signature tiff(ByteBuffer buffer) {
        int width = 0, height = 0;
        A: while (buffer.hasRemaining()) {
            int t1 = buffer.getInt();
            if (t1 == 0 || t1 >= buffer.limit()) break;
            buffer.position(t1);
            // Number of tags in IFD
            int n = buffer.getShort();
            for (int i = 0; i < n; i++) {
                // Tag identifying code
                int tag = buffer.getShort();
                if (tag == 0x100) {
                    buffer.position(buffer.position() + 6);
                    width = buffer.getInt();
                } else if (tag == 0x101) {
                    buffer.position(buffer.position() + 6);
                    height = buffer.getInt();
                    break A;
                } else buffer.position(buffer.position() + 10);
            }
        }
        return new Signature("tiff", "image/tiff", width, height);
    }

    public static class Signature {
        public int width, height;
        public String contentType, extension;

        @Override
        public String toString() {
            return extension + "(" + contentType + "): " + width + "x" + height;
        }

        public Signature(String extension, String contentType) {
            this(extension, contentType, 0, 0);
        }

        public Signature(String extension, String contentType, int width, int height) {
            this.extension = extension;
            this.contentType = contentType;
            this.width = width;
            this.height = height;
        }
    }
}
