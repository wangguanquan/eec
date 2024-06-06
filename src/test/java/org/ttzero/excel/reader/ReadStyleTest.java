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


package org.ttzero.excel.reader;

import org.junit.Test;
import org.ttzero.excel.entity.style.Border;
import org.ttzero.excel.entity.style.BorderStyle;
import org.ttzero.excel.entity.style.Fill;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;
import org.ttzero.excel.entity.style.NumFmt;
import org.ttzero.excel.entity.style.PatternType;
import org.ttzero.excel.entity.style.Styles;

import java.awt.Color;
import java.io.IOException;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.ttzero.excel.reader.ExcelReaderTest.testResourceRoot;

/**
 * @author guanquan.wang at 2023-01-03 20:57
 */
public class ReadStyleTest {
    @Test public void testReadStyles() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("4.xlsx"))) {
            reader.sheet(0).rows().forEach(row -> {
                for (int i = row.fc; i < row.lc; i++) {
                    String rc = new String(org.ttzero.excel.entity.Sheet.int2Col(i + 1)) + row.getRowNum();
                    Cell cell = row.getCell(i);
                    Styles styles = row.getStyles();
                    int style = row.getCellStyle(cell);
                    switch (rc) {
                        case "A1": {
                            Font font = styles.getFont(style);
                            assertTrue(font.isBold());
                            assertEquals(font.getSize(), 11);
                            assertEquals(font.getName(), "微软雅黑");

                            Fill fill = styles.getFill(style);
                            assertEquals(fill.getFgColor(), new Color(102, 102, 153));

                            assertEquals("center", Horizontals.of(styles.getHorizontal(style)));
                            break;
                        }
                        case "B4": {
                            Font font = styles.getFont(style);
                            assertEquals(font.getName(), "Cascadia Mono");
                            assertEquals(font.getSize(), 24);
                            assertEquals("right", Horizontals.of(styles.getHorizontal(style)));

                            Border border = styles.getBorder(style);
                            Border.SubBorder leftBorder = border.getBorderLeft();
                            assertEquals(leftBorder.getStyle(), BorderStyle.HAIR);
                            assertEquals(leftBorder.getColor(), Color.RED);

                            Border.SubBorder bottomBorder = border.getBorderBottom();
                            assertEquals(bottomBorder.getStyle(), BorderStyle.DOUBLE);
                            assertEquals(bottomBorder.getColor(), Color.BLACK);
                            break;
                        }
                        case "E7": {
                            Font font = styles.getFont(style);
                            assertTrue(font.isBold());
                            assertTrue(font.isItalic());
                            assertEquals(font.getSize(), 36);
                            assertEquals(font.getName(), "Consolas");

                            assertEquals("left", Horizontals.of(styles.getHorizontal(style)));

                            Fill fill = styles.getFill(style);
                            assertEquals(fill.getPatternType(), PatternType.gray125);
                            assertEquals(fill.getBgColor(), new Color(123, 193, 203));
                            break;
                        }
                        case "F10": {
                            NumFmt fmt = styles.getNumFmt(style);
                            assertEquals(fmt.getCode(), "d-mmm-yy");

                            Font font = styles.getFont(style);
                            assertTrue(font.isStrikeThru());
                            break;
                        }
                    }
                }
            });
        }
    }

    @Test public void testSpecialIndexedColor() throws IOException {
        try (ExcelReader reader = ExcelReader.read(testResourceRoot().resolve("#145.xlsx"))) {
            Styles styles = reader.getStyles();
            Border border = styles.getBorder(styles.getStyleByIndex(2));

            // indexed color = 10 (255, 0, 0)
            // 实际颜色 (170, 170, 170)

            Border.SubBorder borderLeft = border.getBorderLeft();
            assertEquals(borderLeft.style, BorderStyle.THIN);
            assertEquals(borderLeft.color, new Color(170, 170, 170));

            Border.SubBorder borderTop = border.getBorderTop();
            assertEquals(borderTop.style, BorderStyle.THIN);
            assertEquals(borderTop.color, new Color(170, 170, 170));

            Border.SubBorder borderRight = border.getBorderRight();
            assertEquals(borderRight.style, BorderStyle.THIN);
            assertEquals(borderRight.color, new Color(170, 170, 170));

            Border.SubBorder borderBottom = border.getBorderBottom();
            assertEquals(borderBottom.style, BorderStyle.THIN);
            assertEquals(borderBottom.color, new Color(170, 170, 170));
        }
    }

}
