package net.cua.excel.processor;

/**
 * Int值转其它有意义值
 * 一般用于将状态值，或者枚举值转换为用户感知的具有实际意义的值
 * Created by guanquan.wang on 2017/10/13.
 */
@FunctionalInterface
public interface IntConversionProcessor {
    /**
     * Int值包括byte, char, short, int
     * @param n 数据库值或原对象值
     * @return 转换后的值
     */
    Object conversion(int n);
}
