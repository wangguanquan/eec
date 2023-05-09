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

import java.io.IOException;
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
    public static Map<String, String> whitelist = new HashMap<String, String>() {{
        put("png", null);
        put("jpeg", null);
        put("gif", null);
        put("tiff", null);
        put("bmp", null);
        put("ico", "image/x-ico");
        put("tif", "image/tiff");
        put("emf", "image/x-emf");
        put("wmf", "image/x-wmf");
        put("webp", "image/png");
    }};
    private FileSignatures() { }

    public static Signature test(Path path) {
        try (InputStream is = Files.newInputStream(path)) {
            byte[] bytes = new byte[1 << 9];
            int n = is.read(bytes);
            return test(ByteBuffer.wrap(bytes, 0, n));
        } catch (IOException ex) {

        }
        return null;
    }

    public static Signature test(ByteBuffer buffer) {
        if (buffer.remaining() < 64) return null;
        int t0 = buffer.getShort() & 0xFFFF;

        String extension = null;
        int width = 0, height = 0;

        // Maybe jpeg
        if (t0 == 0xFFD8) {
            extension = "jpeg";
            for (; buffer.remaining() >= 4; ) {
                int t1 = buffer.getShort() & 0xFFFF, n = buffer.getShort() & 0xFFFF;
                if (t1 == 0xFFC0) {
                    if (buffer.remaining() >= 5) {
                        buffer.get();
                        height = buffer.getShort() & 0xFFFF;
                        width = buffer.getShort() & 0xFFFF;
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
            width = buffer.getInt();
            height = buffer.getInt();
        }

        if (extension != null) {
            String contentType = whitelist.get(extension);
            return new Signature(extension, contentType != null ? contentType : "image/" + extension, width, height);
        }

        buffer.position(0);
        t0 = buffer.getInt();

        switch (t0) {
            // Maybe PNG
            case 0x89504E47:
                int t1 = buffer.getInt();
                if (t1 == 0x0D0A1A0A) {
                    extension = "png";
                    buffer.getLong();
                    width = buffer.getInt();
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
                width = buffer.getShort() & 0xFFFF;
                height = buffer.getShort() & 0xFFFF;
                break;
            // ICO
            case 0x100:
                extension = "ico"; break; // TODO
            // EMF
            case 0x01000000:
                extension = "emf"; break; // TODO
            // WMF
            case 0xD7CDC69A:
            case 0x01000900:
                extension = "wmf"; break; // TODO
            // WEBP
            case 0x52494646:
                extension = "webp"; break; // TODO
            default:
        }

        if (extension != null) {
            String contentType = whitelist.getOrDefault(extension, "image/" + extension);
            return new Signature(extension, contentType, width, height);
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

    // <Default Extension="png" ContentType="image/png"/>
    // <Default Extension="svg" ContentType="image/unknown"/>
    // <Default Extension="emf" ContentType="image/x-emf"/>
    // <Default Extension="jpeg" ContentType="image/jpeg"/>
    // <Default Extension="wmf" ContentType="image/x-wmf"/>
    // <Default Extension="gif" ContentType="image/gif"/>
    // <Default Extension="psd" ContentType="image/vnd.adobe.photoshop"/>
    // <Default Extension="tif" ContentType="image/tiff"/>
    // <Default Extension="tiff" ContentType="image/tiff"/>
    // <Default Extension="bmp" ContentType="image/bmp"/>
    // ico 可以换后缀直接保存
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

    public static boolean isOpenXMLSupportExtension(Signature signature) {
        return false;
    }

}
