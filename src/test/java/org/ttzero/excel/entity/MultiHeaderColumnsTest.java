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


package org.ttzero.excel.entity;

import org.junit.Test;
import org.ttzero.excel.annotation.ExcelColumn;
import org.ttzero.excel.annotation.HeaderComment;
import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.entity.style.Font;
import org.ttzero.excel.entity.style.Horizontals;

import java.awt.Color;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;


/**
 * @author guanquan.wang at 2022-06-27 23:24
 */
public class MultiHeaderColumnsTest extends SQLWorkbookTest {
    @Test public void testRepeatAnnotations() throws IOException {
        List<RepeatableEntry> list = RepeatableEntry.randomTestData();
        new Workbook().setWaterMark(WaterMark.of("勿外传"))
            .setAutoSize(true)
            .addSheet(new ListSheet<>(list))
            .writeTo(defaultTestPath.resolve("Repeat Columns Annotation.xlsx"));

        try (ExcelReader reader = ExcelReader.read(defaultTestPath.resolve("Repeat Columns Annotation.xlsx"))) {
            List<RepeatableEntry> readList = reader.sheet(0).header(1, 4).bind(RepeatableEntry.class).rows()
                .map(row -> (RepeatableEntry) row.get()).collect(Collectors.toList());

            assert list.size() == readList.size();
            for (int i = 0, len = list.size(); i < len; i++)
                assert list.get(i).equals(readList.get(i));
        }
    }

    @Test public void testPagingRepeatAnnotations() throws IOException {
        new Workbook("Repeat Paging Columns Annotation").setAutoSize(true)
            .addSheet(new ListSheet<>(RepeatableEntry.randomTestData(10000)).setSheetWriter(new XMLWorksheetWriter() {
                @Override
                public int getRowLimit() {
                    return 500;
                }
            })).writeTo(defaultTestPath);
    }

    @Test public void testMultiOrderColumnSpecifyOnColumn() throws IOException {
        new Workbook("Multi specify columns 2").setAutoSize(true)
            .addSheet(new ListSheet<>("期末成绩", ListObjectSheetTest.Student.randomTestData()
                , new Column("共用表头").addSubColumn(new Column("学号", "id"))
                , new Column("共用表头").addSubColumn(new Column("姓名", "name"))
                , new Column("成绩", "score") {
                @Override
                public int getHeaderStyleIndex() {
                    return styles.of(styles.addFont(this.getFont()) | Horizontals.CENTER);
                }

                public Font getFont() {
                    return new Font("宋体", 12, Color.RED).bold();
                }
            }
            )).writeTo(defaultTestPath);
    }

    @Test public void testMultiOrderColumnSpecifyOnColumn3() throws IOException {
        new Workbook("Multi specify columns 3").setAutoSize(true)
            .addSheet(new ListSheet<>("期末成绩", ListObjectSheetTest.Student.randomTestData()
                , new Column().addSubColumn(new ListSheet.EntryColumn("共用表头")).addSubColumn(new Column("学号", "id").setHeaderComment(new Comment("abc", "content")))
                , new ListSheet.EntryColumn("共用表头").addSubColumn(new Column("姓名", "name"))
                , new Column("成绩", "score")
            )).writeTo(defaultTestPath);
    }

    @Test public void testResultSet() throws SQLException, IOException {
        try (Connection con = getConnection()) {
            new Workbook("Multi ResultSet columns 2", author).setAutoSize(true)
                .setConnection(con)
                .addSheet("select id, name, age, create_date, update_date from student order by age"
                    , new Column("通用").setHeaderStyle(794694).addSubColumn(new Column("学号", int.class))
                    , new Column("通用").addSubColumn(new Column("性名", String.class))
                    , new Column("通用").addSubColumn(new Column("年龄", int.class).setHeaderStyle(794691))
                    , new Column("创建时间", Timestamp.class)
                    , new Column("更新", Timestamp.class).setColIndex(1)
                )
                .writeTo(defaultTestPath);
        }
    }

    public static final String[] provinces = {"江苏省", "湖北省", "浙江省", "广东省"};
    public static final String[][] cities = {{"南京市", "苏州市", "无锡市", "徐州市"}
        , {"武汉市", "黄冈市", "黄石市", "孝感市", "宜昌市"}
        , {"杭州市", "温州市", "绍兴市", "嘉兴市"}
        , {"广州市", "深圳市", "佛山市"}
    };
    public static final String[][][] areas = {{
        {"玄武区", "秦淮区", "鼓楼区", "雨花台区", "栖霞区"}
        , {"虎丘区", "吴中区", "相城区", "姑苏区", "吴江区"}
        , {"锡山区", "惠山区", "滨湖区", "新吴区", "江阴市"}
        , {"鼓楼区", "云龙区", "贾汪区", "泉山区"}
    }, {
        {"江岸区", "江汉区", "硚口区", "汉阳区", "武昌区", "青山区", "洪山区", "东西湖区"}
        , {"黄州区", "团风县", "红安县"}
        , {"黄石港区", "西塞山区", "下陆区", "铁山区"}
        , {"孝南区", "孝昌县", "大悟县", "云梦县"}
        , {"西陵区", "伍家岗区", "点军区"}
    }, {
        {"上城区", "下城区", "江干区", "拱墅区", "西湖区", "滨江区", "余杭区", "萧山区"}
        , {"鹿城区", "龙湾区", "洞头区"}
        , {"越城区", "柯桥区", "上虞区", "新昌县", "诸暨市", "嵊州市"}
        , {"南湖区", "秀洲区", "嘉善县", "海盐县", "海宁市", "平湖市", "桐乡市"}
    }, {
        {"荔湾区", "白云区", "天河区", "黄埔区", "番禺区", "花都区"}
        , {"罗湖区", "福田区", "南山区", "龙岗区"}
        , {"禅城区", "南海区", "顺德区", "三水区", "高明区"}
    }};

    public static class RepeatableEntry {
        @ExcelColumn("TOP")
        @ExcelColumn("K")
        @ExcelColumn
        @ExcelColumn("订单号")
        private String orderNo;
        @ExcelColumn("TOP")
        @ExcelColumn("K")
        @ExcelColumn("A")
        @ExcelColumn("收件人")
        private String recipient;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("A")
        @ExcelColumn("省")
        private String province;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("A")
        @ExcelColumn("市")
        private String city;
        @ExcelColumn("TOP")
        @ExcelColumn("收件地址")
        @ExcelColumn("B")
        @ExcelColumn("区")
        private String area;
        @ExcelColumn("TOP")
        @ExcelColumn(value = "收件地址", comment = @HeaderComment("精确到门牌号"))
        @ExcelColumn("B")
        @ExcelColumn("详细地址")
        private String detail;

        public RepeatableEntry() {}

        public RepeatableEntry(String orderNo, String recipient, String province, String city, String area, String detail) {
            this.orderNo = orderNo;
            this.recipient = recipient;
            this.province = province;
            this.city = city;
            this.area = area;
            this.detail = detail;
        }

        public static List<RepeatableEntry> randomTestData(int n) {
            List<RepeatableEntry> list = new ArrayList<>(n);
            for (int i = 0, p, c; i < n; i++) {
                list.add(new RepeatableEntry(Integer.toString(Math.abs(random.nextInt())), getRandomString(8) + 2, provinces[p = random.nextInt(provinces.length)], cities[p][c = random.nextInt(cities[p].length)], areas[p][c][random.nextInt(areas[p][c].length)], "xx街" + (random.nextInt(10) + 1) + "号"));
            }
            return list;
        }

        public static List<RepeatableEntry> randomTestData() {
            int n = random.nextInt(100) + 1;
            return randomTestData(n);
        }

        public String getOrderNo() {
            return orderNo;
        }

        public String getRecipient() {
            return recipient;
        }

        public String getProvince() {
            return province;
        }

        public String getCity() {
            return city;
        }

        public String getArea() {
            return area;
        }

        public String getDetail() {
            return detail;
        }

        @Override
        public int hashCode() {
            return orderNo.hashCode();
        }

        @Override
        public boolean equals(Object o) {
            if (o instanceof RepeatableEntry) {
                RepeatableEntry other = (RepeatableEntry) o;
                return Objects.equals(orderNo, other.orderNo)
                    && Objects.equals(recipient, other.recipient)
                    && Objects.equals(province, other.province)
                    && Objects.equals(city, other.city)
                    && Objects.equals(detail, other.detail);
            }
            return false;
        }

        @Override
        public String toString() {
            return orderNo + " | " + recipient + " | " + province + " | " + city + " | " + area + " | " + detail;
        }
    }
}
