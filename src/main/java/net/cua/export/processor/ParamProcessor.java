package net.cua.export.processor;

import java.sql.PreparedStatement;
import java.sql.SQLException;

/**
 * 设置SQL参数用
 * @author wanggq
 */
@FunctionalInterface
public interface ParamProcessor {

    void build(PreparedStatement ps) throws SQLException;
}
