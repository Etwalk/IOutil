package annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 按照行列的具体位置来解析excel的
 * @author Sherlock.Wu
 * @date 2019/11/25
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD, ElementType.FIELD})
public @interface ExcelCellBeanAnnotation {
    /**
     *该数据在那一行
     * @return
     */
    int row() ;

    /**
     * 该数据在那一列
     * @return
     */
    int index();
}
