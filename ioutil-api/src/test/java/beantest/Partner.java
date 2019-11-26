package beantest;

import annotation.ExcelCellListAnnotation;
import annotation.ExcelCellListBeginRowAnnotation;
import annotation.ExcelCellLocationAnnotation;
import lombok.Data;
import lombok.ToString;

/**
 * @author Sherlock.Wu
 * @date 2019/11/19
 */
@Data
@ToString
@ExcelCellListBeginRowAnnotation(beginRow =11 )
public class Partner {
    @ExcelCellListAnnotation(index = 0)
    private String name;
    @ExcelCellListAnnotation(index = 1)
    private String address;
    @ExcelCellListAnnotation(index = 2)
    private String buildUpDate;
    @ExcelCellListAnnotation(index = 3)
    private String paymentDays;
}
