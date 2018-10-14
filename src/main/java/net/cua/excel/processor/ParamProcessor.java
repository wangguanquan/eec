package net.cua.excel.processor;

import java.sql.PreparedStatement;
import java.sql.SQLException;

/**
 * 设置SQL参数用
 * Created by guanquan.wang at 2017/9/13.
 */
@FunctionalInterface
public interface ParamProcessor {
    /**
     * 设置SQL参数
     * @param ps PreparedStatement
     * @throws SQLException
     */
    void build(PreparedStatement ps) throws SQLException;
}
