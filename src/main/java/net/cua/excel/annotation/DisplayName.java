package net.cua.excel.annotation;

import java.lang.annotation.*;

/**
 * 导出对象数组时指定对象的导出信息
 * Created by guanquan.wang at 2018-01-30 13:23
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface DisplayName {
    /**
     * header cell
     * @return
     */
    String value() default "";

    /**
     * share body string
     * @return
     */
    boolean share() default false;

    /**
     * skip column when read excel
     * @return
     */
    boolean skip() default false;
}
