package net.cua.excel.util;

import org.junit.Test;

import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

import static net.cua.excel.Print.println;

/**
 * Create by guanquan.wang at 2019-02-11 11:14
 */
public class DateUtilTest {
    private long t0 = 118_751_670_000_000_000L;
    private long t1 = 121_105_206_000_000_000L;
    @Test public void testBiffToTimestamp() {
        Timestamp timestamp = DateUtil.Biff.toTimestamp(t0);
        println(timestamp);

        timestamp = DateUtil.Biff.toTimestamp(t1);
        println(timestamp);

    }

    @Test public void testBiffLocalDateTime() {
        LocalDateTime ldt = DateUtil.Biff.toLocalDateTime(t0);
        println(ldt);

        ldt = DateUtil.Biff.toLocalDateTime(t1);
        println(ldt);
    }

    @Test public void testBiffLocalDate() {
        LocalDate ld = DateUtil.Biff.toLocalDate(t0);
        println(ld);

        ld = DateUtil.Biff.toLocalDate(t1);
        println(ld);
    }

    @Test public void testBiffToTime() {
        Time time = DateUtil.Biff.toTime(t0);
        println(time);

        time = DateUtil.Biff.toTime(t1);
        println(time);

        time = DateUtil.Biff.toTime(DateUtil.Biff.toDateTimeValue(LocalDateTime.now()));
        println(time);
    }

    @Test public void testToDateTimeValue() {
        LocalDateTime ldt = DateUtil.Biff.toLocalDateTime(t0);
        assert t0 == DateUtil.Biff.toDateTimeValue(ldt);

        ldt = DateUtil.Biff.toLocalDateTime(t1);
        assert t1 == DateUtil.Biff.toDateTimeValue(ldt);

        Timestamp ts0 = new Timestamp(System.currentTimeMillis());
        long t = DateUtil.Biff.toDateTimeValue(ts0);
        Timestamp ts1 = DateUtil.Biff.toTimestamp(t);

        assert ts0.equals(ts1);

        Date now = new Date();
        t = DateUtil.Biff.toDateTimeValue(now);
        assert now.equals(DateUtil.Biff.toDate(t));
    }

}
