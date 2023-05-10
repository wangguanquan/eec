package org.ttzero.excel.util;

import org.junit.Test;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;

import static org.junit.Assert.assertEquals;
import static org.ttzero.excel.util.DateUtil.DAYS_1900_TO_1970;

/**
 *
 *
 * @author CarlJia
 * @date 2023-03-27
 */
public class DateUtilTest {

    @Test
    public void toLocalDate() {
        LocalDate localDate = DateUtil.toLocalDate(24014);
        LocalDate expectedLocalDate = LocalDate.of(1965, 9, 29);
        assertEquals(expectedLocalDate, localDate);

        LocalDate localDate1 = DateUtil.toLocalDate(-25567 + DAYS_1900_TO_1970);
        LocalDate expectedLocalDate1 = LocalDate.of(1900, 1, 1);
        assertEquals(expectedLocalDate1, localDate1);

        LocalDate ts = DateUtil.toLocalDate(-86890 + DAYS_1900_TO_1970);
        assertEquals(ts, LocalDate.of(1732, 2, 8));
    }

    @Test
    public void toLocalDateTime() {
        LocalDateTime localDateTime = DateUtil.toLocalDateTime(24014);
        LocalDateTime expectedLocalDateTime = LocalDateTime.of(1965, 9, 29, 0, 0, 0);
        assertEquals(expectedLocalDateTime, localDateTime);

        LocalDateTime localDateTime1 = DateUtil.toLocalDateTime(44927.00000579861);
        assertEquals(localDateTime1, LocalDateTime.of(2023, 1, 1, 0, 0, 1));

        LocalDateTime localDateTime2 = DateUtil.toLocalDateTime(44927.000011516204);
        assertEquals(localDateTime2, LocalDateTime.of(2023, 1, 1, 0, 0, 1));

        LocalDateTime localDateTime3 = DateUtil.toLocalDateTime(44927.000011516204);
        assertEquals(localDateTime3, LocalDateTime.of(2023, 1, 1, 0, 0, 1));

        LocalDateTime localDateTime4 = DateUtil.toLocalDateTime(44727.99998842592);
        assertEquals(localDateTime4, LocalDateTime.of(2022, 6, 15, 23, 59, 59));

        LocalDateTime localDateTime5 = DateUtil.toLocalDateTime(44728.99998836806);
        assertEquals(localDateTime5, LocalDateTime.of(2022, 6, 16, 23, 59, 59));
    }

    @Test public void toTimestamp() {
        Timestamp ts = DateUtil.toTimestamp(-86890 + DAYS_1900_TO_1970);
        assertEquals(ts, Timestamp.valueOf("1732-02-08 00:00:00"));

        Timestamp ts2 = DateUtil.toTimestamp(44927.000011516204);
        assertEquals(ts2, Timestamp.valueOf("2023-01-01 00:00:01"));

        Timestamp ts3 = DateUtil.toTimestamp(44728.99998836806);
        assertEquals(ts3, Timestamp.valueOf("2022-06-16 23:59:59"));

        Timestamp ts4 = DateUtil.toTimestamp(2);
        assertEquals(ts4, Timestamp.valueOf("1900-01-01 00:00:00"));

        Timestamp ts5 = DateUtil.toTimestamp("2023-01-01");
        assertEquals(ts5, Timestamp.valueOf("2023-01-01 00:00:00"));

        Timestamp ts6 = DateUtil.toTimestamp("2023-01-01 1");
        assertEquals(ts6, Timestamp.valueOf("2023-01-01 01:00:00"));

        Timestamp ts7 = DateUtil.toTimestamp("2023-01-01 1:");
        assertEquals(ts7, Timestamp.valueOf("2023-01-01 01:00:00"));

        Timestamp ts8 = DateUtil.toTimestamp("2023-01-01 1:2");
        assertEquals(ts8, Timestamp.valueOf("2023-01-01 01:02:00"));

        Timestamp ts9 = DateUtil.toTimestamp("2023-01-01 1:2:");
        assertEquals(ts9, Timestamp.valueOf("2023-01-01 01:02:00"));

        Timestamp ts10 = DateUtil.toTimestamp("2023-01-01 1:2:3");
        assertEquals(ts10, Timestamp.valueOf("2023-01-01 01:02:03"));
    }
}