package maintest;

import api.ParseExcel;
import beantest.ContractInfoBean;
import beantest.Partner;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ParaseExcelTestMain {
    public static void main(String[] args) throws Exception{
//        URL resource = ParaseExcelTestMain.class.getClassLoader().getResource("testparase.xlsx");
//        String path = resource.getPath();
        InputStream inputStream = ParaseExcelTestMain.class.getResourceAsStream("/customerInfo.xls");
//        System.out.println(pathone);
//        String path =pathone+"testparase.xlsx";
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len;
        byte[] dataBytes;
        while ((len = inputStream.read(buffer)) != -1 ) {
            baos.write(buffer, 0, len);
        }
        baos.flush();
        dataBytes = baos.toByteArray();


        ParseExcel contractInfoBeanExcelUtil = new ParseExcel();
        List contractInfoBeans = contractInfoBeanExcelUtil.readExcelBeans(new ByteArrayInputStream(dataBytes), Partner.class,0);
        Object obj = ParseExcel.parseExcelToObject(new ByteArrayInputStream(dataBytes),ContractInfoBean.class,0);
        System.out.println(obj);
        if(ContractInfoBean.class.isInstance(obj)){
            ContractInfoBean contractInfoBean = (ContractInfoBean)obj;
            System.out.println(contractInfoBean);
        }

    }
}
