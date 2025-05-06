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

package org.ttzero.excel.reader;

import static org.ttzero.excel.entity.Sheet.int2Col;
import static org.ttzero.excel.entity.Sheet.toCoordinate;
import static org.ttzero.excel.reader.ExcelReader.coordinateToLong;
import static org.ttzero.excel.util.ExtBufferedWriter.getChars;
import static org.ttzero.excel.util.ExtBufferedWriter.stringSize;

/**
 * 范围，它包含起始到结束行列值，应用于合并单元格时指定单元格范围和指定工作表的有效范围，
 *
 * <p>Excel的列由{@code A-Z}大写字母组成，行由{@code 1,2,3}数字组成，每个坐标都由列+行组成
 * {@code 1}行{@code 1}列表示为{@code A1}以此类推，当列到达{@code Z}之后就由两位字母联合组成，
 * {@code Z}的下一列表示为{@code AA}，同理{@code ZZ}列的下一列表示为{@code AAA}</p>
 *
 *
 * <p>范围值包含两个坐标，例{@code A1:B5}它表示从1行1列到5行2列的范围，如果起始坐标和结束坐标一样
 * 也就是压缩到一个单元格可以简写为起始坐标{@code A1:A1}被记为{@code A1}</p>
 *
 * @author guanquan.wang at 2019-12-20 10:07
 */
public class Dimension {
    /**
     * 起始行号 (one base)
     */
    public final int firstRow;
    /**
     * 末尾行号 (one base)
     */
    public final int lastRow;
    /**
     * 起始列号 (one base)
     */
    public final short firstColumn;
    /**
     * 末尾列号 (one base)
     */
    public final short lastColumn;
    /**
     * 宽 = 末尾列-起始列+1
     * 高 = 末尾行-起始行+1
     */
    public final int width, height;

    /**
     * 创建一个范围值， 起始和结束相同
     *
     * @param firstRow    起始行号 (one base)
     * @param firstColumn 起始列号 (one base)
     */
    public Dimension(int firstRow, short firstColumn) {
        this(firstRow, firstColumn, firstRow, firstColumn);
    }

    /**
     * 创建一个范围值
     *
     * @param firstRow    起始行号 (one base)
     * @param firstColumn 起始列号 (one base)
     * @param lastRow     末尾行号 (one base)
     * @param lastColumn  末尾列号 (one base)
     */
    public Dimension(int firstRow, short firstColumn, int lastRow, short lastColumn) {
        this.firstRow = Math.max(firstRow, 1);
        this.firstColumn = (short) Math.max(firstColumn, 1);
        this.lastRow = lastRow > 0 ? lastRow : this.firstRow;
        this.lastColumn = lastColumn > 0 ? lastColumn : this.firstColumn;

        this.width = this.lastColumn - this.firstColumn + 1;
        this.height = this.lastRow - this.firstRow + 1;
        if (width < 1 || height < 1)
            throw new IllegalArgumentException("Dimension(firstRow:" + firstRow + ",firstColumn:" + firstColumn + ",lastRow=" + lastRow + ",lastColumn=" + lastColumn + ") contains invalid range");
    }

    /**
     * 解析范围字符串，有效的范围字符串至少包含一个起始坐标，最多包含两个坐标，坐标间使用{@code ‘:’}分隔
     *
     * @param range 范围字符串 像{@code A2:B2}
     * @return the {@link Dimension} entry
     */
    public static Dimension of(String range) {
        int i = range.indexOf(':');

        long f = 0L, t = 0L;
        if (i < 0) {
            f = coordinateToLong(range);
        } else if (i == 0) {
            t = coordinateToLong(range.substring(i + 1));
        } else {
            f = coordinateToLong(range.substring(0, i));
            t = coordinateToLong(range.substring(i + 1));
        }
        return new Dimension((int) (f >> 16), (short) f, (int) (t >> 16), (short) t);
    }

    /**
     * 获取起始行号，最小为1
     *
     * @return the first row number
     */
    public int getFirstRow() {
        return firstRow;
    }

    /**
     * 获取末尾行号
     *
     * @return the last row number
     */
    public int getLastRow() {
        return lastRow;
    }

    /**
     * 获取起始列号，最小为1
     *
     * @return the first column number
     */
    public short getFirstColumn() {
        return firstColumn;
    }

    /**
     * 获取末尾列号
     *
     * @return the last column number
     */
    public short getLastColumn() {
        return lastColumn;
    }

    /**
     * 获取范围宽度，宽 = 末尾列-起始列+1
     *
     * @return 宽度
     */
    public int getWidth() {
        return width;
    }

    /**
     * 获取范围高度，高 = 末尾行-起始行+1
     *
     * @return 高度
     */
    public int getHeight() {
        return height;
    }

    @Override
    public String toString() {
        return toCoordinate(firstRow, firstColumn)
            + (lastRow > firstRow || lastColumn > firstColumn ? ":" + toCoordinate(lastRow, lastColumn) : "");
    }

    /**
     * 转为引用字符串
     *
     * @return 引用字符串
     */
    public String toReferer() {
        char[] chars;
        if (lastRow > firstRow || lastColumn > firstColumn) {
            int i = 0, c0 = firstColumn <= 26 ? 1 : firstColumn <= 702 ? 2 : 3 , r0 = stringSize(firstRow), c1 = lastColumn <= 26 ? 1 : lastColumn <= 702 ? 2 : 3, r1 = stringSize(lastRow);
            chars = new char[c0 + r0 + c1 + r1 + 5];
            chars[i++] = '$';
            System.arraycopy(int2Col(firstColumn), 0, chars, i, c0);
            i += c0;
            chars[i++] = '$';
            getChars(firstRow, i += r0, chars);
            chars[i++] = ':';
            chars[i++] = '$';
            System.arraycopy(int2Col(lastColumn), 0, chars, i, c1);
            i += c1;
            chars[i++] = '$';
            getChars(lastRow, i + r1, chars);
        } else {
            int c0 = firstColumn <= 26 ? 1 : firstColumn <= 702 ? 2 : 3;
            chars = new char[c0 + stringSize(firstRow) + 2];
            chars[0] = '$';
            System.arraycopy(int2Col(firstColumn), 0, chars, 1, c0);
            chars[c0 + 1] = '$';
            getChars(firstRow, chars.length, chars);
        }
        return new String(chars);
    }

    /**
     * 检查指定坐标是否在本范围中
     *
     * @param r 行号
     * @param c 列号
     * @return true 如果坐标在范围中
     */
    public boolean checkRange(int r, int c) {
        return r >= firstRow && r <= lastRow && c >= firstColumn && c <= lastColumn;
    }

    @Override
    public int hashCode() {
        return ((firstColumn << 24) | lastColumn) ^ ((firstRow << 24) | lastRow);
    }

    @Override
    public boolean equals(Object o) {
        boolean r = this == o;
        if (!r && o instanceof Dimension) {
            Dimension other = (Dimension) o;
            r = other.firstRow == firstRow && other.firstColumn == firstColumn
                    && other.lastRow == lastRow && other.lastColumn == lastColumn;
        }
        return r;
    }
}
