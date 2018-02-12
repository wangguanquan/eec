package net.cua.export;

import net.cua.export.entity.e7.style.DefaultNumFmt;
import net.cua.export.entity.e7.style.Horizontals;
import net.cua.export.entity.e7.style.Styles;
import org.junit.Test;

import java.util.Arrays;

/**
 * Created by guanquan.wang on 2017/10/20.
 */
public class TesSytles {
//    @Before public void init() {
//        styles = new Workbook().getStyles();
//    }
    @Test public void unpack() {
        System.out.println(
                Arrays.toString(Styles.unpack(Styles.defaultStringBorderStyle()))
        );
        System.out.println(
                Arrays.toString(Styles.unpack(Styles.clearHorizontal(Styles.defaultStringBorderStyle())))
        );
        System.out.println(
                Arrays.toString(Styles.unpack(Styles.clearHorizontal(Styles.defaultStringBorderStyle()) | Horizontals.CENTER))
        );
    }

    @Test public void defaultNumFmt() {
        System.out.println(DefaultNumFmt.get("yyyy\"5E74\"m\"6708\"d\"65E5\""));
    }
}
