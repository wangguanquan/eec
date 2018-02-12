package net.cua.export.processor;

import java.sql.PreparedStatement;
import java.sql.SQLException;

/**
 * 使用实体做SQL参数
 * @param <T> T
 * @author guanquan.wang
 */
@FunctionalInterface
public interface BeanParamProcessor<T> {

    void build(PreparedStatement ps, T object) throws SQLException;
}
