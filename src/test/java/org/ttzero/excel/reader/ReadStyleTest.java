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
                    if ("A1".equals(rc)) {
                        Font font = styles.getFont(style);
                        assert font.isBold();
                        assert font.getSize() == 11;
                        assert font.getName().equals("微软雅黑");

                        Fill fill = styles.getFill(style);
                        assert fill.getFgColor().equals(new Color(102, 102, 153));

                        assert "center".equals(Horizontals.of(styles.getHorizontal(style)));
                    } else if ("B4".equals(rc)) {
                        Font font = styles.getFont(style);
                        assert font.getName().equals("Cascadia Mono");
                        assert font.getSize() == 24;
                        assert "right".equals(Horizontals.of(styles.getHorizontal(style)));

                        Border border = styles.getBorder(style);
                        Border.SubBorder leftBorder = border.getBorderLeft();
                        assert leftBorder.getStyle() == BorderStyle.HAIR;
                        assert leftBorder.getColor().equals(Color.RED);

                        Border.SubBorder bottomBorder = border.getBorderBottom();
                        assert bottomBorder.getStyle() == BorderStyle.DOUBLE;
                        assert bottomBorder.getColor().equals(Color.BLACK);
                    } else if ("E7".equals(rc)) {
                        Font font = styles.getFont(style);
                        assert font.isBold();
                        assert font.isItalic();
                        assert font.getSize() == 36;
                        assert font.getName().equals("Consolas");

                        assert "left".equals(Horizontals.of(styles.getHorizontal(style)));

                        Fill fill = styles.getFill(style);
                        assert fill.getPatternType() == PatternType.gray125;
                        assert fill.getBgColor().equals(new Color(123, 193, 203));
                    } else if ("F10".equals(rc)) {
                        NumFmt fmt = styles.getNumFmt(style);
                        assert fmt.getCode().equals("d-mmm-yy");

                        Font font = styles.getFont(style);
                        assert font.isStrikeThru();
                    }
                }
            });
        }
    }
}
