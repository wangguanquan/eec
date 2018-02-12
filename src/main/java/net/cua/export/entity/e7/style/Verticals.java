package net.cua.export.entity.e7.style;

/**
 * Created by guanquan.wang at 2018-02-11 14:59
 */
public class Verticals {
    public static final int CENTER = 0 // Align Center
            , BOTTOM = 1 << Styles.INDEX_VERTICAL // Align Bottom
            , TOP    = 2 << Styles.INDEX_VERTICAL // Align Top
            , BOTH   = 3 << Styles.INDEX_VERTICAL // Vertical Justification
            ;

    private static final String[] _names = {"center", "bottom", "top", "both"};

    public static String of(int n) {
        return _names[n];
    }
}
