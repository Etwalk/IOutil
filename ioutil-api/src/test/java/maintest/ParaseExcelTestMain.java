package maintest;

import api.ExcelBuilderUtil;
import api.ParaseExcel;
import beantest.ContractInfoBean;
import sun.security.util.Resources;

import java.io.InputStream;
import java.net.URL;
import java.util.List;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ParaseExcelTestMain {
    public static void main(String[] args) {
        URL resource = ParaseExcelTestMain.class.getClassLoader().getResource("testparase.xlsx");
        String path = resource.getPath();
//        InputStream inputStream = ParaseExcelTestMain.class.getResourceAsStream("/testparase.xlsx");
//        System.out.println(pathone);
//        String path =pathone+"testparase.xlsx";


        ParaseExcel<ContractInfoBean> contractInfoBeanExcelUtil = new ParaseExcel<ContractInfoBean>();
        List<ContractInfoBean> contractInfoBeans = contractInfoBeanExcelUtil.readExcel(path,ContractInfoBean.class);
        System.out.println(contractInfoBeans);
        ContractInfoBean contractInfoBean = new ContractInfoBean();
        System.out.println(contractInfoBean);

        ExcelBuilderUtil<ContractInfoBean> excelBuilderUtil = new ExcelBuilderUtil<ContractInfoBean>();
        excelBuilderUtil.exportExcelToFile("客户信息表",contractInfoBeans,"D:/test.xls",ContractInfoBean.class);
    }
}
