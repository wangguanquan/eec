package net.cua.excel.processor;

import net.cua.excel.entity.e7.style.Styles;

/**
 * 样式转换器
 * Created by guanquan.wang at 2017/10/13.
 */
@FunctionalInterface
public interface StyleProcessor {
    /**
     * 样式转换器
     * 添加样式时必须使用sst.add方法添加，然后将返回的int值做为转换器的返回值
     * eg:
     * <pre><code lang='java'>
     *    StyleProcessor sp = (o, style, sst) // 将背景改为黄色
     *      -> style |= Styles.clearFill(style) | sst.addFill(new Fill(Color.yellow));
     * </code></pre>
     * @param o 当前单元格值
     * @param style 当前单元格样式
     * @param sst 样式类，整个Workbook共享样式
     * @return 新样式
     */
    int build(Object o, int style, Styles sst);
}
