package org.ttzero.excel.util;

import org.junit.Test;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;

import static java.time.ZoneOffset.UTC;
import static org.junit.Assert.*;
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
    }

    @Test
    public void toLocalDateTime() {
        LocalDateTime localDateTime = DateUtil.toLocalDateTime(24014);
        LocalDateTime expectedLocalDateTime = LocalDateTime.of(1965, 9, 29, 0, 0, 0);
        assertEquals(expectedLocalDateTime, localDateTime);
    }
}