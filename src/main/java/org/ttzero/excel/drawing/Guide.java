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


package org.ttzero.excel.drawing;

/**
 * @author guanquan.wang at 2023-07-28 10:53
 */
public class Guide {
    public String name;
    public String fmla;

    public Guide() { }

    public Guide(String name, String fmla) {
        this.name = name;
        this.fmla = fmla;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getFmla() {
        return fmla;
    }

    public void setFmla(String fmla) {
        this.fmla = fmla;
    }
}
