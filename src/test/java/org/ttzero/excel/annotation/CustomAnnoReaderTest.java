/*
 * Copyright (c) 2017-2022, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.annotation;

import org.junit.Test;
import org.ttzero.excel.customAnno.ExcelProperty;
import org.ttzero.excel.entity.ListSheet;
import org.ttzero.excel.reader.ExcelReadException;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.HeaderRow;
import org.ttzero.excel.reader.Sheet;
import org.ttzero.excel.reader.XMLRow;
import org.ttzero.excel.reader.XMLSheet;
import org.ttzero.excel.util.DateUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Collectors;

import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;
import static org.ttzero.excel.util.StringUtil.EMPTY;

/**
 * @author guanquan.wang at 2022-05-03 18:34
 */
public class CustomAnnoReaderTest {

    @Test public void testCustomeAnno() throws IOException {
        try (ExcelReader reader = new MyExcelReader(testResourceRoot().resolve("1.xlsx"))) {
            List<Entry> list = reader.sheets().flatMap(Sheet::dataRows).map(row -> row.to(Entry.class)).collect(Collectors.toList());
            assert "4 | 3 | XuSu2gFg32 | 2018-11-21 | true | F | false".equals(list.get(0).toString());
            assert "5 | 5 | 7OZXtUuk | 2018-11-21 | true | P | false".equals(list.get(60).toString());
            assert "3 | 2 | Ae9CNO6eTu | 2018-11-21 | true | B | false".equals(list.get(93).toString());
        }
    }

    public static class MyExcelReader extends ExcelReader {
        public MyExcelReader(Path path) throws IOException {
            super(path);
        }

        @Override
        protected XMLSheet sheetFactory(int option) {
            return new XMLSheet() {
                @Override
                public XMLRow createRow() {
                    return new MyRow();
                }
            };
        }
    }

    public static class GameConverter implements Converter<Integer> {
        final String[] names = {"未知", "LOL", "WOW", "极品飞车", "守望先锋", "怪物世界"};

        @Override
        public Integer convertToJavaData(CellData cellData) throws Exception {
            return StringUtil.indexOf(names, cellData.getStringValue());
        }

        @Override
        public CellData convertToExcelData(Integer value) throws Exception {
            return value >= 0 && value < names.length ? new CellData<>(names[value]) : new CellData<>(names[0]);
        }
    }

    public static class MyRow extends XMLRow {
        @Override
        public HeaderRow asHeader() {
            return new HeaderRow() {
                @Override
                protected ListSheet.EntryColumn createColumn(AccessibleObject ao) {
                    ExcelProperty ep = ao.getAnnotation(ExcelProperty.class);
                    if (ep != null) {
                        Class<? extends Converter> converterClazz = ep.converter();
                        ListSheet.EntryColumn column;
                        if (converterClazz != Converter.AutoConverter.class) {
                            ConvertColumn cc = new ConvertColumn(ep.value()[0], converterClazz);

                            try {
                                cc.converter = converterClazz.newInstance();
                            } catch (IllegalAccessException | InstantiationException e1) {
                                LOGGER.warn("无法解析Converter: {}", converterClazz);
                            }
                            column = cc;
                        } else {
                            column = new ListSheet.EntryColumn(ep.value()[0]);
                        }
                        return column;
                    }
                    // Row Num
                    RowNum rowNum = ao.getAnnotation(RowNum.class);
                    if (rowNum != null) {
                        return new ListSheet.EntryColumn(EMPTY, RowNum.class);
                    }
                    return null;
                }

                @Override
                protected void methodPut(int i, org.ttzero.excel.reader.Row row, Object t) throws IllegalAccessException, InvocationTargetException {
                    Class<?> fieldClazz = columns[i].clazz;

                    // 兼容处理
                    if (Converter.class.isAssignableFrom(fieldClazz)) {
                        convert((ConvertColumn) columns[i], row, t);
                    } else {
                        super.methodPut(i, row, t);
                    }
                }

                @Override
                protected void fieldPut(int i, org.ttzero.excel.reader.Row row, Object t) throws IllegalAccessException {
                    Class<?> fieldClazz = columns[i].clazz;

                    // 兼容处理
                    if (Converter.class.isAssignableFrom(fieldClazz)) {
                        convert((ConvertColumn) columns[i], row, t);
                    } else {
                        super.fieldPut(i, row, t);
                    }
                }

                private void convert(ConvertColumn column, org.ttzero.excel.reader.Row row, Object t) {
                    Converter<?> converter = column.converter;
                    if (converter != null) {
                        try {
                            Object o = converter.convertToJavaData(toCellData(row, column.colIndex));
                            if (column.method != null) {
                                column.method.invoke(t, o);
                            } else {
                                column.field.set(t, o);
                            }
                        } catch (Exception e) {
                            throw new ExcelReadException(e);
                        }
                    }
                }

                private CellData toCellData(org.ttzero.excel.reader.Row row, int i) {
                    CellData cellData = new CellData();
                    switch (row.getCellType(i)) {
                        case STRING: cellData.setStringValue(row.getString(i)); break;
                        case INTEGER:
                        case LONG: cellData.setNumberValue(row.getDecimal(i)); break;
                        case BOOLEAN: cellData.setBooleanValue(row.getBoolean(i)); break;
                        default:
                            cellData.setData(row.getString(i)); break;
                    }
                    return cellData;
                }

            }.with(this);
        }
    }

    public static class Entry {
        @RowNum
        private int row;
        @ExcelProperty("渠道ID")
        private Integer channelId;
        @ExcelProperty(value = "游戏", converter = GameConverter.class)
        private Integer gameCode;
        @ExcelProperty
        private String account;
        @ExcelProperty("注册时间")
        private java.util.Date registered;
        @ExcelProperty("是否满30级")
        private boolean up30;
        @ExcelProperty("敏感信息不导出")
        private int id; // not export
        private String address;
        @ExcelProperty("VIP")
        private char c;

        private boolean vip;

        public boolean isUp30() {
            return up30;
        }

        public void setC(char c) {
            this.c = c;
            this.vip = c == 'A';
        }

        public boolean isVip() {
            return vip;
        }

        @Override
        public String toString() {
            return channelId + " | "
                    + gameCode + " | "
                    + account + " | "
                    + (registered != null ? DateUtil.toDateString(registered) : null) + " | "
                    + up30 + " | "
                    + c + " | "
                    + isVip()
                    ;
        }
    }

    public static class ConvertColumn extends ListSheet.EntryColumn {
        private Converter<?> converter;

        public ConvertColumn() {
        }

        public ConvertColumn(String name) {
            super(name);
        }

        public ConvertColumn(String name, Class<?> clazz) {
            super(name, clazz);
        }

        public Converter<?> getConverter() {
            return converter;
        }

        public void setConverter(Converter<?> converter) {
            this.converter = converter;
        }
    }
}
