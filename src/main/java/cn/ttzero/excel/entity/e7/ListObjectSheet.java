/*
 * Copyright (c) 2009, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.entity.e7;

import cn.ttzero.excel.entity.ExportException;
import cn.ttzero.excel.util.ExtBufferedWriter;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;
import cn.ttzero.excel.annotation.DisplayName;
import cn.ttzero.excel.annotation.NotExport;
import cn.ttzero.excel.entity.WaterMark;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by guanquan.wang at 2018-01-26 14:48
 */
public class ListObjectSheet<T> extends Sheet {
    private List<T> data;
    private Field[] fields;

    public ListObjectSheet(Workbook workbook) {
        super(workbook);
    }

    public ListObjectSheet(Workbook workbook, String name, Column[] columns) {
        super(workbook, name, columns);
    }

    public ListObjectSheet(Workbook workbook, String name, WaterMark waterMark, Column[] columns) {
        super(workbook, name, waterMark, columns);
    }

    @Override
    public void close() {
        data.clear();
        data = null;
    }

    public ListObjectSheet<T> setData(final List<T> data) {
        this.data = data;
        return this;
    }

    @Override
    public void writeTo(Path xl) throws IOException, ExportException {
        Path worksheets = xl.resolve("worksheets");
        if (!Files.exists(worksheets)) {
            FileUtil.mkdir(worksheets);
        }
        String name = getFileName();
        workbook.what("0010", getName());

        // create header columns
        fields = init();
        if (fields.length == 0 || fields[0] == null) {
            workbook.what("Worksheet do't has columns.");
            return;
        }

        File sheetFile = worksheets.resolve(name).toFile();

        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            // Write header
            writeBefore(bw);
            // Main data
            // Write sheet data
            if (getAutoSize() == 1) {
                for (T o : data) {
                    writeRowAutoSize(o, bw);
                }
            } else {
                for (T o : data) {
                    writeRow(o, bw);
                }
            }

            // Write foot
            writeAfter(bw);

        } catch (IllegalAccessException e) {
            throw new ExportException(e);
        }

        // resize columns
        boolean resize = false;
        for (Column hc : columns) {
            if (hc.getWidth() > 0.000001) {
                resize = true;
                break;
            }
        }

        if (getAutoSize() == 1 || resize) {
            autoColumnSize(sheetFile);
        }

        // relationship
        relManager.write(worksheets, name);
    }

    private static final String[] exclude = {"serialVersionUID", "this$0"};
    protected Field[] init() {
        Object o = workbook.getFirst(data);
        if (columns == null || columns.length == 0) {
            Field[] fields = o.getClass().getDeclaredFields();
            List<Column> list = new ArrayList<>(fields.length);
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                String gs = field.toGenericString();
                NotExport notExport = field.getAnnotation(NotExport.class);
                if (notExport != null || StringUtil.indexOf(exclude, gs.substring(gs.lastIndexOf('.') + 1)) >= 0) {
                    fields[i] = null;
                    continue;
                }
                DisplayName dn = field.getAnnotation(DisplayName.class);
                if (dn != null && StringUtil.isNotEmpty(dn.value())) {
                    list.add(new Column(dn.value(), field.getName(), field.getType()).setShare(dn.share()));
                } else {
                    list.add(new Column(field.getName(), field.getName(), field.getType()).setShare(dn != null && dn.share()));
                }
            }
            columns = new Column[list.size()];
            list.toArray(columns);
            for (int i = 0; i < columns.length; i++) {
                columns[i].setSst(workbook.getStyles());
            }
            // clear not export fields
            for (int len = fields.length, n = len - 1; n >= 0; n--) {
                if (fields[n] != null) {
                    fields[n].setAccessible(true);
                    continue;
                }
                if (n < len - 1) {
                    System.arraycopy(fields, n + 1, fields, n, len - n - 1);
                }
                len--;
            }
            return fields;
        } else {
            Field[] fields = new Field[columns.length];
            Class<?> clazz = o.getClass();
            for (int i = 0; i< columns.length; i++) {
                Column hc = columns[i];
                try {
                    fields[i] = clazz.getDeclaredField(hc.getKey());
                    fields[i].setAccessible(true);
                    if (hc.getClazz() == null) {
                        hc.setClazz(fields[i].getType());
//                        DisplayName dn = field.getAnnotation(DisplayName.class);
//                        if (dn != null) {
//                            hc.setShare(hc.isShare() || dn.share());
//                            if (StringUtil.isEmpty(hc.getName())
//                                    && StringUtil.isNotEmpty(dn.value())) {
//                                hc.setName(dn.value());
//                            }
//                        }
                    }
                } catch (NoSuchFieldException e) {
                    throw new ExportException("Column " + hc.getName() + " not declare in class " + clazz);
                }
            }
            return fields;
        }

    }

    /**
     * write object
     *
     * @param o
     * @param bw
     * @throws IOException
     */
    protected void writeRowAutoSize(T o, ExtBufferedWriter bw) throws IOException, IllegalAccessException {
        if (o == null) {
            writeEmptyRow(bw);
            return;
        }
        int r = ++rows;
        // logging
        if (r % 1_0000 == 0) {
            workbook.what("0014", String.valueOf(r));
        }
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Field field;
        Object e;
        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            field = fields[i];
            Class<?> clazz = hc.getClazz();
            // t n=numeric (default), s=string, b=boolean
            if (isString(clazz)) {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeStringAutoSize(bw, s, i);
            }
            else if (isDate(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.util.Date date = (java.util.Date) e;
                    writeDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isDateTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    Timestamp ts = (Timestamp) e;
                    writeTimestamp(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isChar(clazz)) {
                Character c = (Character) field.get(o);
                if (c != null) writeCharAutoSize(bw, c, i);
                else writeNull(bw, i);
            }
            else if (isInt(clazz)) {
                Integer n = (Integer) field.get(o);
                if (n != null) writeIntAutoSize(bw, n, i);
                else writeNull(bw, i);
            }
            else if (isLong(clazz)) {
                Long l = (Long) field.get(o);
                if (l != null) writeLong(bw, l, i);
                else writeNull(bw, i);
            }
            else if (isFloat(clazz)) {
                Double d = (Double) field.get(o);
                if (d != null) writeDouble(bw, d, i);
                else writeNull(bw, i);
            }
            else if (isBool(clazz)) {
                Boolean bool = (Boolean) field.get(o);
                if (bool != null) writeBoolean(bw, bool, i);
                else writeNull(bw, i);
            }
            else if (isBigDecimal(clazz)) {
                e = field.get(o);
                if (e != null) {
                    BigDecimal bd = (BigDecimal) e;
                    writeBigDecimal(bw, bd, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isLocalDate(clazz)) {
                e = field.get(o);
                if (e != null) {
                    LocalDate date = (java.time.LocalDate) e;
                    writeLocalDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isLocalDateTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    LocalDateTime ts = (java.time.LocalDateTime) e;
                    writeLocalDateTime(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.sql.Time t = (java.sql.Time) e;
                    writeTime(bw, t, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isLocalTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.time.LocalTime t = (java.time.LocalTime) e;
                    writeLocalTime(bw, t, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeStringAutoSize(bw, s, i);
            }
        }
        bw.write("</row>");
    }

    protected void writeRow(T o, ExtBufferedWriter bw) throws IOException, IllegalAccessException {
        if (o == null) {
            writeEmptyRow(bw);
            return;
        }
        // Row number
        int r = ++rows;
        // logging
        if (r % 1_0000 == 0) {
            workbook.what("0014", String.valueOf(r));
        }
        final int len = columns.length;
        bw.write("<row r=\"");
        bw.writeInt(r);
        bw.write("\" spans=\"1:");
        bw.writeInt(len);
        bw.write("\">");

        Field field;
        Object e;
        for (int i = 0; i < len; i++) {
            Column hc = columns[i];
            field = fields[i];
            Class<?> clazz = hc.getClazz();
            // t n=numeric (default), s=string, b=boolean
            if (isString(clazz)) {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeString(bw, s, i);
            }
            else if (isDate(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.util.Date date = (java.util.Date) e;
                    writeDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isDateTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    Timestamp ts = (Timestamp) e;
                    writeTimestamp(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isChar(clazz)) {
                Character c = (Character) field.get(o);
                if (c != null) writeChar(bw, c, i);
                else writeNull(bw, i);
            }
            else if (isInt(clazz)) {
                Integer n = (Integer) field.get(o);
                if (n != null) writeInt(bw, n, i);
                else writeNull(bw, i);
            }
            else if (isLong(clazz)) {
                Long l = (Long) field.get(o);
                if (l != null) writeLong(bw, l, i);
                else writeNull(bw, i);
            }
            else if (isFloat(clazz)) {
                Double d = (Double) field.get(o);
                if (d != null) writeDouble(bw, d, i);
                else writeNull(bw, i);
            }
            else if (isBool(clazz)) {
                Boolean bool = (Boolean) field.get(o);
                if (bool != null) writeBoolean(bw, bool, i);
                else writeNull(bw, i);
            }
            else if (isBigDecimal(clazz)) {
                e = field.get(o);
                if (e != null) {
                    BigDecimal bd = (BigDecimal) e;
                    writeBigDecimal(bw, bd, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isLocalDate(clazz)) {
                e = field.get(o);
                if (e != null) {
                    LocalDate date = (java.time.LocalDate) e;
                    writeLocalDate(bw, date, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isLocalDateTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    LocalDateTime ts = (java.time.LocalDateTime) e;
                    writeLocalDateTime(bw, ts, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.sql.Time t = (java.sql.Time) e;
                    writeTime(bw, t, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else if (isLocalTime(clazz)) {
                e = field.get(o);
                if (e != null) {
                    java.time.LocalTime t = (java.time.LocalTime) e;
                    writeLocalTime(bw, t, i);
                } else {
                    writeNull(bw, i);
                }
            }
            else {
                String s = (e = field.get(o)) != null ? e.toString() : null;
                writeString(bw, s, i);
            }
        }
        bw.write("</row>");
    }

}
