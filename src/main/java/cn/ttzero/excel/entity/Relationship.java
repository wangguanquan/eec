/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
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

package cn.ttzero.excel.entity;

import java.io.Serializable;

/**
 * Created by guanquan.wang on 2017/9/30.
 */
public class Relationship implements Serializable, Cloneable {

	private static final long serialVersionUID = 1L;
	private String target;
    private String type;
    private String id;

    public Relationship() {}
    public Relationship(String target, String type) {
        this.target = target;
        this.type = type;
    }
    public Relationship(String id, String target, String type) {
        this.id = id;
        this.target = target;
        this.type = type;
    }

    public String getTarget() {
        return target;
    }

    public void setTarget(String target) {
        this.target = target;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    @Override
    public Relationship clone() {
        Relationship r;
        try {
            r = (Relationship) super.clone();
        } catch (CloneNotSupportedException e) {
            r = new Relationship();
            r.id = id;
            r.target = target;
            r.type = type;
        }
        return r;
    }
}
