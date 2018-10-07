package net.cua.excel.entity.e7;

import java.io.Serializable;

/**
 * Created by guanquan.wang on 2017/9/30.
 */
public class Relationship implements Serializable, Cloneable {
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
