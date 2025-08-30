/*
 * Copyright (c) 2017-2025, guanquan.wang@yandex.com All Rights Reserved.
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


package org.ttzero.excel.util;

import org.dom4j.io.SAXReader;
import org.xml.sax.SAXException;

import java.lang.reflect.Method;

/**
 * 低版本SAXReader兼容处理
 *
 * @author wangguanquan3 at 2025-08-30 16:06
 */
public class SAXReaderUtil {

    /**
     * 创建SAXReader并增加安全相关feature配置
     *
     * @return 安全的 {@link SAXReader}实例
     */
    public static SAXReader createDefault() {
        SAXReader reader;
        try {
            Method method = SAXReader.class.getDeclaredMethod("createDefault");
            reader = (SAXReader) method.invoke(null);
        } catch (Exception ex) {
            reader = new SAXReader();
            try {
                reader.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
                reader.setFeature("http://xml.org/sax/features/external-general-entities", false);
                reader.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            } catch (SAXException e) {
                // nothing to do, incompatible reader
            }
        }
        return reader;
    }
}
