package org.ttzero.excel.annotation;

import java.lang.annotation.*;

/**
 * Excel对象注解
 * <p>
 * 当ExcelEntity对象中有嵌套项时，需要使用此注解
 * </p>
 * @author Chai at 2025/4/7
 */
@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface NestedObject {

    /**
     * 嵌套对象中的列名格式化规则，采用{@link String#format(String, Object...)}方式格式化内嵌对象中的列名
     */
    String columnNameFormat() default "%s";

    /**
     * 嵌套对象内部的列索引开始值
     * <p>
     *     当同一个类中出现多个相同类型的内嵌对象时，可以通过此值来设置之后的列索引的起始值
     * </p>
     */
    int startColIndex() default -1;

    /**
     * 此嵌套对象中的列名是否继承此外层对象的列名格式化规则
     * <p>
     *     例：A 类中包含 B 类，B 类中包含 C 类，B 类的 {@link NestedObject#columnNameFormat()} 为 "B_%s"，C 类的 {@link NestedObject#columnNameFormat()} 为 "C_%s"，
     *     当B类中C属性加入此注解且 columnNameFormatExtend 为 true 时，C 类的列名将继承 B 类的 columnNameFormat 规则，
     *     则C类中的列名规则实际为："B_C_%s"
     *     (默认会去除上层嵌套对象{@link NestedObject#columnNameFormat()}中的%s，只会保留最低层级中的%s占位符)
     * </p>
     */
    boolean columnNameFormatExtend() default false;
}
