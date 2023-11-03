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

package org.ttzero.excel.entity;

import java.io.Serializable;

/**
 * 资源的关联关系，Excel将部分资源统一存放在一个公共区域，例如图片，图表，批注等。
 * 工作表添加图片，图表的时候只需要添加一个关联就可以引用到这些资源，这样做可以起到共享资源的目的。
 *
 * <p>每个{@code Relationship}都包含一个{@code Id}值，其它引用者通过该Id值确定唯一的源，
 * {@code Target}保存资源的相对位置，这里的位置是相对于引用者而不是当前关联表存放位置，
 * 最后一个{@code Type}属性是一个固定的{@code schema}值表示资源的类型，这里{@link org.ttzero.excel.manager.Const.Relationship}
 * 定义了当前支持的{@code schema}</p>
 *
 * <p>{@code Relationship}并不会独立存在，它对应一个{@code rels}管理者，{@link org.ttzero.excel.manager.RelManager}
 * 就是关系管理器的角色，每个引用者都包含一个独立的关系管理器，所以每个关系管理器中{@code Relationship}的{@code Id}值都是从{@code 1}开始的，
 * 即使相同的资源在不同管理器中的{@code Id}值都可能不同</p>
 *
 * @author guanquan.wang on 2017/9/30.
 */
public class Relationship implements Serializable, Cloneable {

    private static final long serialVersionUID = 1L;
    /**
     * 资源的相对位置，这里的位置是相对于引用者而不是当前关联表存放位置
     */
    private String target;
    /**
     * 资源的一个固定{@code schema}值表示资源的类型
     */
    private String type;
    /**
     * 引用者通过该Id查询对应的引用资源
     */
    private String id;

    public Relationship() { }

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
