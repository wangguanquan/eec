package net.cua.excel.processor;

/**
 * 监听操作日志
 * Create by guanquan.wang at 2018-10-13
 */
@FunctionalInterface
public interface Watch {
    /**
     * 监听
     * @param msg 输出信息
     */
    void what(String msg);
}
