package net.cua.export.processor;

import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * 暴露SQL结果集
 * @author wanggq
 */
@FunctionalInterface
public interface ResultSetProcessor {

    void process(ResultSet resultSet) throws SQLException;
}
