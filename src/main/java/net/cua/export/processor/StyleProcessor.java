package net.cua.export.processor;

/**
 * Created by wanggq on 2017/10/13.
 */
@FunctionalInterface
public interface StyleProcessor {
    int build(Object o, int style);
}
