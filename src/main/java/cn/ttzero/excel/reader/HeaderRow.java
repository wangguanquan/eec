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

package cn.ttzero.excel.reader;

import cn.ttzero.excel.annotation.DisplayName;
import cn.ttzero.excel.annotation.NotImport;
import cn.ttzero.excel.util.StringUtil;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.StringJoiner;

/**
 * Create by guanquan.wang at 2019-04-17 11:55
 */
public class HeaderRow extends Row {

    String[] names;
    Class<?> clazz;
    Field[] fields;
    int[] columns;
    Class<?>[] fieldClazz;
    Object t;

    private HeaderRow() { }

    HeaderRow(SharedString sst, int startRow) {
        super(sst, startRow);
    }

    static HeaderRow with(Row row) {
        HeaderRow hr = new HeaderRow();
        hr.names = new String[row.p2];
        Cell c;
        for (int i = row.p1 - 1; i < row.p2; i++) {
            c = row.cells[i];
            // header type is string
            if (c.getT() == 's') {
                hr.names[i] = row.sst.get(c.getNv());
            } else {
                hr.names[i] = c.getSv();
            }
        }
        return hr;
    }

    final boolean is(Class<?> clazz) {
        return this.clazz != null && this.clazz == clazz;
    }

    /**
     * mapping
     * @param clazz the type of binding
     * @return the header row
     */
    final HeaderRow setClass(Class<?> clazz) {
        this.clazz = clazz;
        Field[] fields = clazz.getDeclaredFields();
        int[] index = new int[fields.length];
        int count = 0;
        for (int i = 0, n; i < fields.length; i++) {
            Field f = fields[i];
            // skip not import fields
            NotImport nit = f.getAnnotation(NotImport.class);
            if (nit != null) {
                fields[i] = null;
                continue;
            }
            // field has display name
            DisplayName ano = f.getAnnotation(DisplayName.class);
            if (ano != null && StringUtil.isNotEmpty(ano.value())) {
                n = StringUtil.indexOf(names, ano.value());
                if (n == -1) {
                    logger.warn(clazz + " field [" + ano.value() + "] can't find in header" + Arrays.toString(names));
                    fields[i] = null;
                    continue;
                }
            }
            // no annotation or annotation value is null
            else {
                String name = f.getName();
                n = StringUtil.indexOf(names, name);
                if (n == -1 && (n = StringUtil.indexOf(names, StringUtil.toPascalCase(name))) == -1) {
                    fields[i] = null;
                    continue;
                }
            }

            index[i] = n;
            count++;
        }

        this.fields = new Field[count];
        this.columns = new int[count];
        this.fieldClazz = new Class<?>[count];

        for (int i = fields.length - 1; i >= 0; i--) {
            if (fields[i] != null) {
                count--;
                this.fields[count] = fields[i];
                this.fields[count].setAccessible(true);
                this.columns[count] = index[i];
                this.fieldClazz[count] = fields[i].getType();
            }
        }
        return this;
    }

    /**
     * mapping and instance
     * @param clazz the type of binding
     * @return the header row
     * @throws IllegalAccessException -
     * @throws InstantiationException -
     */
    final HeaderRow setClassOnce(Class<?> clazz) throws IllegalAccessException, InstantiationException {
        setClass(clazz);
        this.t = clazz.newInstance();
        return this;
    }

    final Field[] getFields() {
        return fields;
    }

    final int[] getColumns() {
        return columns;
    }

    final Class<?>[] getFieldClazz() {
        return fieldClazz;
    }

    @SuppressWarnings("unchecked")
    final <T> T getT() {
        return (T) t;
    }

    /**
     * Get T value by column index
     * @param columnIndex the cell index
     * @return T
     */
    @SuppressWarnings("unchecked")
    @Override public String get(int columnIndex) {
        rangeCheck(columnIndex);
        return names[columnIndex];
    }

    @Override public String toString() {
        StringJoiner joiner = new StringJoiner(" | ");
        int i = 0;
        for (; names[i++] == null; );
        for (; i < names.length; i++) {
            joiner.add(names[i]);
        }
        return joiner.toString();
    }

}
