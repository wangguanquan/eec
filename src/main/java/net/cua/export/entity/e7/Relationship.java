package net.cua.export.entity.e7;

/**
 * Created by wanggq on 2017/9/30.
 */
public class Relationship {
    private String target;
    private String type;
    private String id;

    public Relationship() {}
    public Relationship(String target, String type) {
        this.target = target;
        this.type = type;
//        this.id = id;
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
}
