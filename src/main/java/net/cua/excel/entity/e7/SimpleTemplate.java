package net.cua.excel.entity.e7;

import java.nio.file.Path;

/**
 * Created by guanquan.wang at 2018-02-23 17:19
 */
public class SimpleTemplate extends AbstractTemplate {

    public SimpleTemplate(Path zipPath, Workbook wb) {
        super(zipPath, wb);
    }

    @Override
    protected boolean isPlaceholder(String txt) {
        int len = txt.length();
        return len > 3 &&  txt.charAt(0) == '$' && txt.charAt(1) == '{' && txt.charAt(len - 1) == '}';
    }

    private String getKey(String txt) {
        return txt.substring(2, txt.length() - 1).trim();
    }

    @Override
    protected String getValue(String txt) {
        if (map == null) return txt;
        String value, key = getKey(txt);
        if (map.containsKey(key)) {
            value = map.get(key);
        } else value = txt;

        return value;
    }

}
