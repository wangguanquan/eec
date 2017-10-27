package net.cua.export.processor;

import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * ResultSet转实体
 * @param <T> T
 * @author wanggq
 */
@FunctionalInterface
public interface BeanResultSetProcessor<T> {

    T process(ResultSet resultSet) throws SQLException;
}
