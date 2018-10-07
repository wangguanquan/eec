package net.cua.excel.processor;

import java.sql.PreparedStatement;
import java.sql.SQLException;

/**
 * 设置SQL参数用
 * @author guanquan.wang
 */
@FunctionalInterface
public interface ParamProcessor {

    void build(PreparedStatement ps) throws SQLException;
}
