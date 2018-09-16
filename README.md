# eec介绍
Excel导出工具类，目前仅支持xlsx格式的表格导出，表格默认有黑色边框可修改。
1. 支持Statement, ResultSet数据库导出，导出行数无上限，如果数据量超过单个sheet上限会自动分页。
2. 支持 对象数组 和 Map数组 导出。
3. 可以为某列设置阀值高亮显示。如导出学生成绩时低于60分的单元格背景标黄显示。
4. int类型(byte,char,short,int)方便转可被识别文字
5. 设置单元格隐藏
6. 设置列宽自动调节
7. 设置水印（文字，本地＆网络图片）

#### 使用方法
导入eec.jar即可使用

```
git clone https://www.github.com/wangguanquan/ecc.git
mvn install
```

pom.xml添加


```
<dependency>
    <groupId>net.cua</groupId>
    <artifactId>excel-export</artifactId>
    <version>1.0</version>
</dependency>
```

### 以下是部分功能测试代码，更多使用方法请参考test/各测试类
1. 无参SQL固定宽度导出测试,固定宽度20,也可以使用setWidth(int)来重置列宽

```
    @Test public void t1() {
        try (Connection con = dataSource.getConnection()) {
            new Workbook("单表－单Sheet-固定宽度", creator) // 指定workbook名，作者
                .setConnection(con) // 配置数据库连接
                .setAutoSize(true) // 列宽自动调节
                .addSheet("用户注册"
                    , "select id,pro_id,channel_no,aid,account,regist_time,uid,platform_type from wh_regist limit 10"
                    , new Sheet.Column("ID", int.class)
                    , new Sheet.Column("产品ID", int.class)
                    , new Sheet.Column("渠道ID", int.class)
                    , new Sheet.Column("AID", int.class)
                    , new Sheet.Column("注册账号", String.class)
                    , new Sheet.Column("注册时间", Timestamp.class)
                    , new Sheet.Column("CPS用户ID", int.class)
                    , new Sheet.Column("渠道类型", int.class)
                ) // 添加一个sheet页
                .writeTo(defaultPath); // 指定输出位置
        } catch (SQLException | IOException | ExportException e) {
            e.printStackTrace();
        }
    }
```

2. SQL带参数测试，且将满足条件的单元格标红。如果某个列字符串重复率很高时可以将其设为共享达到数据压缩的目的。

```
    @Test public void t2() {
        try (Connection con = dataSource.getConnection();
             FileOutputStream fos = new FileOutputStream(defaultPath.resolve("单页-固定宽度-值转换＆样式转换-输出流.xlsx").toFile())) {
            boolean share = true;
            String[] cs = {"正常", "注销"};
            final Fill fill = new Fill(PatternType.solid, Color.red);
            new Workbook("多Sheet页-固定宽度-值转换＆样式转换", creator)
                    .setConnection(con)
                    .addSheet("CPS渠道列表"
                            , "select id, name, account, status,city from t_user where id between ? and ? and city = ?"
                            , p -> {
                                p.setInt(1, 1);
                                p.setInt(2, 500);
                                p.setString(3, "苏州市");
                            } // 设定SQL参数
                            , new Sheet.Column("编号", int.class)
                            , new Sheet.Column("登录名", String.class) // 登录名都是唯一的可以不设共享
                            , new Sheet.Column("通行证", String.class)
                            , new Sheet.Column("状态", char.class, c -> cs[c], share) // 将0/1用户无感的数字转为文字，并共享字串
                                    .setStyleProcessor((n, style, sst) -> {
                                        if ((int) n == 1) { // 将注销的用户标记
                                            style = Styles.clearFill(style) | sst.addFill(fill); // 注销标红
                                        }
                                        return style;
                                    })
                            , new Sheet.Column("城市", String.class, share) // 共享字串
                    )
                    .addSheet("用户注册"
                            , "select id,pro_id,channel_no,aid,account,regist_time,uid,platform_type from wh_regist limit 10"
                            , new Sheet.Column("ID", int.class)
                            , new Sheet.Column("产品ID", int.class)
                            , new Sheet.Column("渠道ID", int.class)
                            , new Sheet.Column("AID", int.class)
                            , new Sheet.Column("注册账号", String.class)
                            , new Sheet.Column("注册时间", Timestamp.class)
                            , new Sheet.Column("CPS用户ID", int.class)
                            , new Sheet.Column("渠道类型", int.class)
                    )
                    .writeTo(fos); // 输出到output，如果是web导出功能这里可以直接输出到｀response.getOutputStream()｀
        } catch (SQLException | IOException | ExportException e) {
            e.printStackTrace();
        }
    }
```

3. 对象数组 & Map数组支持。对象可以通过注解@DisplayName来设置表头列或共享，使用@NotExport来指定不导出的字段。

```
    /**
     * 测试对象
     */
    public class TestExportEntity {
        @DisplayName("渠道ID")
        private int channelId;
        @DisplayName(value = "游戏", share = true)
        private String pro;
        @DisplayName
        private String account;
        @DisplayName("注册时间")
        private Timestamp registered;
        @DisplayName("是否满30级")
        private boolean up30;
        @NotExport("敏感信息不导出")
        private int id; // not export
        private String address;
        @DisplayName("VIP")
        private char c;
    }

    // Test
    @Test public void t3() {
        long start = System.currentTimeMillis();
        Workbook wb = new Workbook("List<?>-Map<String, ?>-测试", creator)
        .setAutoSize(true) // Auto-size
        .setWaterMark(WaterMark.of("guanquan.wang 2018-01-01")); // 设置水印
            wb.addSheet(new EmptySheet(wb, "空数据"
                    , new Sheet.Column("姓名", String.class)
                    , new Sheet.Column("性别", String.class)
            ).hidden()); // 设置此Sheet为隐藏

            List<Map<String, Object>> mapData = new ArrayList<>();
            for (int i = 0; i < 251; i++) {
                Map<String, Object> map = new HashMap<>();
                map.put("name", "colvin" + i);
                map.put("age", i);
                map.put("date", new Timestamp(start + 12000));
                mapData.add(map);
            }

            // head columns 决定导出顺序
            wb.addSheet("Map测试", mapData
                    , new Sheet.Column("姓名", "name")
                    , new Sheet.Column("年龄", "age")
                        .setStyleProcessor((o, style, sst) -> {
                            if (((int)o) > 150) {
                                style = Styles.clearFill(style) | sst.addFill(Fill.parse("#ff0000"));
                            }
                            return style;
                        })
                    , new Sheet.Column("录入时间", "date")
            );
            // 设置边框
            Border border = new Border();
            border.setBorder(BorderStyle.DOTTED, Color.red);
            border.setBorderBottom(BorderStyle.NONE);
            // 设置填充
            Fill fill = new Fill();
            fill.setPatternType(PatternType.solid);
            fill.setFgColor(Color.GRAY);
            fill.setBgColor(Color.decode("#ccff00"));
            // 设置字体
            Font font = new Font("Klee", 14, Font.Style.bold, Color.white);
            font.setCharset(Charset.GB2312); // 字符集
            wb.getSheet("Map测试").setHeadStyle(font, fill, border); // 重设头部样式

            List<TestExportEntity> objectData = new ArrayList<>();
            int size = random.nextInt(100) + 1;
            String[] proArray = {"LOL", "WOW", "极品飞车", "守望先锋", "怪物世界"};
            TestExportEntity e;
            while (size-->0) {
                e = new TestExportEntity();
                e.id = size;
                e.channelId = random.nextInt(10) + 1;
                e.pro = proArray[random.nextInt(5)];
                e.account = getRandom();
                e.registered = new Timestamp(start += random.nextInt(8000));
                e.up30 = random.nextInt(10) > 3;
                e.address = getRandom();
                e.c = (char) ('A' + random.nextInt(26));
                objectData.add(e);
            }
            wb.addSheet("Object测试", objectData);  //  方式1
            wb.getSheet("Object测试") // Set style
                    .setHeadStyle(Font.parse("'under line' 11 Klee red")
                            , Fill.parse("#666699 solid")
                            , Border.parse("thin #ff0000").setDiagonalDown(BorderStyle.THIN, Color.CYAN));
            wb.setName("New name"); // Rename

            wb.addSheet("Object copy", objectData  // 方式2
                    , new Sheet.Column("渠道ID", "id")
                            .setType(Sheet.Column.TYPE_RMB)
                    , new Sheet.Column("游戏", "pro")
                    , new Sheet.Column("账户", "account")
                    , new Sheet.Column("是否满30级", "up30")
                    , new Sheet.Column("渠道", "channelId", n -> n < 5 ? "自媒体" : "联众", true)
                    , new Sheet.Column("注册时间", "registered")
            );
        try {
            wb.writeTo(defaultPath);
        } catch (IOException | ExportException ex) {
            ex.printStackTrace();
        }
    }
```

4. 有时候你可能会使用模板来规范格式，不固定的部分使用${key}标记，Excel导出时使用Map或者Java bean传入。

如有以下格式模板文件template.xlsx

>                       通知书
>     ${name } 同学，在本次期末考试的成绩是 ${score}。
>                                 ${date }

测试代码

```
    @Test public void t4() {
        try (FileInputStream fis = new FileInputStream(new File(defaultPath.toString(), "template.xlsx"))) {
            // Map data
            Map<String, Object> map = new HashMap<>();
            map.put("name", "guanquan.wang");
            map.put("score", 90);
            map.put("date", "2018-02-12 12:22:29");

            // java bean
//            BindEntity entity = new BindEntity();
//            entity.score = 67;
//            entity.name = "张三";
//            entity.date = new Timestamp(System.currentTimeMillis());

            new Workbook("模板导出", creator)
                    .withTemplate(fis, map) // 绑定模板
                    .writeTo(defaultPath); // 写到某个文件夹
        } catch (IOException | ExportException e) {
            e.printStackTrace();
        }
    }
```

## TODO LIST

1. excel文件增加导出scripts功能
2. list data with template
3. 对excel文件设置密码 (AES-128 encrypted)
4. 多线程支持，多个sheet数据同时写
5. share多线程支持
6. 异常出理
7. 单元格隐藏 -

## BUG

1. new Fill无法设置正确的背景