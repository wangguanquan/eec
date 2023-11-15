package org.ttzero.excel.processor;

/**
 * 转换器，包含{@link #conversion}和{@link #reversion}方法，前者用于输出时将Java数据转为Excel数据，
 * 后者正好相反输入时将Excel数据转为Java数据
 *
 * @author guanquan.wang on 2023-11-14 11:49
 */
public interface Converter<T> extends ConversionProcessor {
    /**
     * 输入转换器，读取Excel时将单元格的值转为指定类型{@code T}
     *
     * @param v Excel原始值
     * @return 转换后的值
     */
    T reversion(String v);

    /**
     * 无类型转换，默认
     */
    final class None implements Converter<Object> {

        @Override
        public Object reversion(String v) {
            return v;
        }

        @Override
        public Object conversion(Object v) {
            return v;
        }
    }
}
