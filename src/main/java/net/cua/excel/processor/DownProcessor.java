package net.cua.excel.processor;

import java.nio.file.Path;

/**
 * Excel定操作完成后可以做后续操作
 * Created by guanquan.wang at 2018/6/13.
 */
@FunctionalInterface
public interface DownProcessor {
    /**
     * 执行此方法
     * @param path excel临时位置
     */
    void exec(Path path);
}
