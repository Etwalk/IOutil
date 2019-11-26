package beantest;

import annotation.BuildExcelAnnotation;
import annotation.ExcelCellBeanAnnotation;
import annotation.ExcelCellListBeginRowAnnotation;
import lombok.Data;
import lombok.ToString;

import java.util.List;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
@Data
@ToString
public class ContractInfoBean {
    @ExcelCellBeanAnnotation(row = 3,index = 1)
    private String typeId;
    @ExcelCellBeanAnnotation(row = 2,index = 3)
    @BuildExcelAnnotation("合同编号")
    private String contracId;
    @ExcelCellBeanAnnotation(row = 2,index = 5)
    @BuildExcelAnnotation("客户姓名")
    private String username;
    @ExcelCellBeanAnnotation(row = 2,index = 1)
    private String type;
    @ExcelCellListBeginRowAnnotation(beginRow = 11)
    List<Partner> partnerList;

}
