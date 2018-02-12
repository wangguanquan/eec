package net.cua.export;

import net.cua.export.tmap.TIntIntHashMap;
import org.junit.Test;

import java.util.Arrays;

/**
 * Created by guanquan.wang on 2017/10/17.
 */
public class TestTrove4j {
    @Test
    public void put() {
        TIntIntHashMap map = new TIntIntHashMap();
        map.put(23, 0);
        map.put(46, 1);
        map.put(-2432, 2);
        map.put(-45, 3);

        System.out.println(map.size());

        int[] values = map.values();
        System.out.println(Arrays.toString(values));

        System.out.println(Arrays.toString(unpackStyle(532547)));
    }

    static final int[] move_left = {24, 18, 12, 6, 3, 0};
    int[] unpackStyle(int style) {
        int[] styles = new int[6];
        styles[0] = style >>> move_left[0];
        styles[1] = style << 8 >>> move_left[1] + 8;
        styles[2] = style << 14 >>> move_left[2] + 14;
        styles[3] = style << 20 >>> move_left[3] + 20;
        styles[4] = style << 26 >>> move_left[4] + 26;
        styles[5] = style << 29 >>> move_left[5] + 29;
        return styles;
    }
}
