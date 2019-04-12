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

import cn.ttzero.excel.manager.Const;
import cn.ttzero.excel.annotation.TopNS;
import cn.ttzero.excel.util.FileUtil;
import cn.ttzero.excel.util.StringUtil;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;

import java.io.IOException;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 字符串共享，一个workbook的所有worksheet共享
 *
 * Created by guanquan.wang at 2017/10/10.
 */
@TopNS(prefix = "", value = "sst", uri = Const.SCHEMA_MAIN)
public class SharedStrings {
    // 存储共享字符
    private Map<String, Integer> elements;
    private int count; // workbook所有字符串(shared属性为true)的个数
    private static final int MAX_CACHE_SIZE = 8192;

    SharedStrings() {
        elements = new LinkedHashMap<>();
    }

    private ThreadLocal<char[]> charCache = ThreadLocal.withInitial(() -> new char[1]);
    public int get(char c) {
        char[] cs = charCache.get();
        cs[0] = c;
        return get(new String(cs));
    }

    /**
     * TODO 每个sheet采用one by one的方式输出，暂不考虑并发
     * FIXME OOM occur
     * @param key the string
     * @return index of the string in the SST
     */
    public int get(String key) {
        Integer n = elements.get(key);
        if (n == null) {
            if (elements.size() < MAX_CACHE_SIZE) {
                elements.put(key, n = elements.size());
            } else {
                return -1;
            }
        }
        count++;
        return n;
    }

    public void write(Path root) throws IOException {
        TopNS topNS = getClass().getAnnotation(TopNS.class);

        DocumentFactory factory = DocumentFactory.getInstance();
        //use the factory to create a root element
        Element rootElement = factory.createElement(topNS.value(), topNS.uri()[0]);
        rootElement.addAttribute("uniqueCount", String.valueOf(elements.size()));
        rootElement.addAttribute("count", String.valueOf(count));

        elements.forEach((k,v) -> rootElement.addElement("si").addElement("t").setText(k));

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDiskNoFormat(doc, root.resolve(StringUtil.lowFirstKey(getClass().getSimpleName() + Const.Suffix.XML))); // write to desk

        // destroy
        destroy();
    }

    /**
     * clear memory
     */
    private void destroy() {
        elements.clear();
        elements = null;
    }
}
