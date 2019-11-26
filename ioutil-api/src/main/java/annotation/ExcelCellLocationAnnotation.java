package annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD, ElementType.FIELD})
/**
 * index值从0开始，如果超过了excel每行的单元格数，则结果为null
 */
public @interface ExcelCellLocationAnnotation  {
    /**
     *该数据在那一行
     * @return
     */
    int row() default -1;

    /**
     * 该数据在那一列
     * @return
     */
    int index() default -1;

    /**
     * 结束的标志行是为空的时候
     * list开始的标志
     * 和row()互相排斥
     * 该数据从那一行开始，到row cell 为空的时候结束
     * @return
     */
    int beginRow() default -1;
}
