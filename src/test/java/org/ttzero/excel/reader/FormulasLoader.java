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

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * 装载formulas预期文件
 *
 * @author guanquan.wang at 2023-11-06 19:05
 */
public class FormulasLoader {
    /**
     * 解析expect文件夹下以$formulas结尾的文件
     *
     * @param path 预期文件路径
     * @return 行列坐标：公式字符串
     */
    public static Map<Long, String> load(Path path) {
        if (!Files.exists(path)) return Collections.emptyMap();
        Map<Long, String> map = new HashMap<>();
        try (BufferedReader reader = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            String line;
            while ((line = reader.readLine()) != null) {
                int i = line.indexOf('=');
                if (i > 0) map.put(ExcelReader.cellRangeToLong(line.substring(0, i)), line.substring(i + 1));
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return map;
    }
}
