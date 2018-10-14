package net.cua.excel.annotation;

import java.lang.annotation.*;

/**
 * 指定字段不导出
 * Created by guanquan.wang at 2018-01-30 15:09
 */
@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface NotExport {
    /**
     * 添加不导出的原因，方便开发者阅读
     * @return 原因
     */
    String value() default "";
}
