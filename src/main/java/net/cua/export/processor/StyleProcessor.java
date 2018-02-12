package net.cua.export.processor;

import net.cua.export.entity.e7.style.Styles;

/**
 * Created by wanggq on 2017/10/13.
 */
@FunctionalInterface
public interface StyleProcessor {
    int build(Object o, int style, Styles sst);
}
