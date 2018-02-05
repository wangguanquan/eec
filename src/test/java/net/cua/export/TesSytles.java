package net.cua.export;

import net.cua.export.entity.e7.Styles;
import org.junit.Test;

import java.util.Arrays;

/**
 * Created by wanggq on 2017/10/20.
 */
public class TesSytles {
//    @Before public void init() {
//        styles = new Workbook().getStyles();
//    }
    @Test public void unpack() {
        System.out.println(
                Arrays.toString(Styles.unpack(Styles.defaultStringStyle()))
        );
        System.out.println(
                Arrays.toString(Styles.unpack(Styles.clearHorizontal(Styles.defaultStringStyle())))
        );
        System.out.println(
                Arrays.toString(Styles.unpack(Styles.clearHorizontal(Styles.defaultStringStyle()) | Styles.Horizontals.CENTER))
        );
    }
}
