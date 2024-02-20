/*
 * Copyright (c) 2017, guanquan.wang@yandex.com All Rights Reserved.
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

package org.ttzero.excel.manager.docProps;

import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.Namespace;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.entity.Storable;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

/**
 * @author guanquan.wang on 2017/9/21.
 */
public abstract class XmlEntity implements Storable {
    @Override
    public void writeTo(Path path) throws IOException {
        TopNS topNs = getClass().getAnnotation(TopNS.class);
        if (topNs == null) throw new IOException(getClass() + " top namespace is required.");
        String[] prefixs = topNs.prefix(), uris = topNs.uri();
        // Use the factory to create a root element
        DocumentFactory factory = DocumentFactory.getInstance();
        Element rootElement = prefixs.length > 0 && StringUtil.isEmpty(prefixs[0]) ? factory.createElement(topNs.value(), uris[0]) : factory.createElement(topNs.value());
        Map<String, Namespace> namespaceMap = new HashMap<>(prefixs.length);
        // Attach namespace
        for (int i = 0; i < prefixs.length; i++) {
            namespaceMap.put(prefixs[i], Namespace.get(prefixs[i], uris[i]));
            rootElement.add(Namespace.get(prefixs[i], uris[i]));
        }

        // 转dom树
        toDom(rootElement, namespaceMap);

        Document doc = factory.createDocument(rootElement);
        FileUtil.writeToDiskNoFormat(doc, path); // write to desk
    }

    /**
     * 对象转Dom树
     *
     * @param root 要节点
     * @param namespaceMap namespace
     */
    abstract void toDom(Element root, Map<String, Namespace> namespaceMap);
}
