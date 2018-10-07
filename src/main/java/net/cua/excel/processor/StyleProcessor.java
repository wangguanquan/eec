package net.cua.excel.processor;

import net.cua.excel.entity.e7.style.Styles;

/**
 * Created by guanquan.wang on 2017/10/13.
 */
@FunctionalInterface
public interface StyleProcessor {
    int build(Object o, int style, Styles sst);
}
