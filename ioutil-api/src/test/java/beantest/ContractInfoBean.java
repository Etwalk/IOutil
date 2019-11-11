package beantest;

import annotation.ExcelBuilderAnnotation;
import annotation.ParaseExcelAnnotation;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ContractInfoBean {
    @ParaseExcelAnnotation(index = 0)
    @ExcelBuilderAnnotation("合同编号")
    private String contracId;
    @ParaseExcelAnnotation(index = 1)
    @ExcelBuilderAnnotation("客户姓名")
    private String username;
    @ParaseExcelAnnotation(index = 2)
    private String type;
    @ParaseExcelAnnotation(index = 3)
    private String typeId;

    public String getContracId() {
        return contracId;
    }
//    @ParaseExcelAnnotation(index = 1)
    public void setContracId(String contracId) {
        this.contracId = contracId;
    }

    public String getUsername() {
        return username;
    }
//    @ParaseExcelAnnotation(index = 2)
    public void setUsername(String username) {
        this.username = username;
    }

    public String getType() {
        return type;
    }
//    @ParaseExcelAnnotation(index = 3)
    public void setType(String type) {
        this.type = type;
    }

    public String getTypeId() {
        return typeId;
    }
//    @ParaseExcelAnnotation(index = 4)
    public void setTypeId(String typeId) {
        this.typeId = typeId;
    }

    @Override
    public String toString() {
        return "ContractInfoBean{" +
                "contracId=" + contracId + '\'' +
                ", username=" + username + '\'' +
                ", type='" + type + '\'' +
                ", typeId='" + typeId + '\'' +
                '}';
    }
}
