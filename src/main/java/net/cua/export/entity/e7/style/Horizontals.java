package net.cua.export.entity.e7.style;

/**
 * Created by wanggq at 2018-02-11 15:02
 */
public class Horizontals {
    // General Horizontal Alignment( Text data is left-aligned. Numbers
    // , dates, and times are right-aligned.Boolean types are centered)
    public static final int GENERAL = 0
            , LEFT = 1 // Left Horizontal Alignment
            , RIGHT = 2 // Right Horizontal Alignment
            , CENTER = 3 // Centered Horizontal Alignment
            , CENTER_CONTINUOUS = 4 // (Center Continuous Horizontal Alignment
            , FILL = 5 // Fill
            , JUSTIFY = 6 // Justify
            , DISTRIBUTED = 7 // Distributed Horizontal Alignment
            ;

    private static final String[] _names = {"general","left","right","center","centerContinuous","fill","justify","distributed"};

    public static String of(int n) {
        return _names[n];
    }
}
