package maintest;

import api.BuildExcelUtil;
import beantest.ContractInfoBean;
import beantest.Partner;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Sherlock.Wu
 * @date 2019/11/26
 */
public class BuilderExcelTestMain {
    public static void main(String[] args) {
//        List<Partner>  list = new ArrayList();
        ContractInfoBean contractInfoBean = new ContractInfoBean();
        contractInfoBean.setContracId("123445555");
        contractInfoBean.setType("身份证");
        contractInfoBean.setTypeId("23eddxssssssddddd");
        contractInfoBean.setUsername("三生");
        List<Partner> partnerList = new ArrayList<>();
        Partner partner = new Partner();
        partner.setAddress("dkkdkdkdkdk");
        partner.setBuildUpDate("ldldkdkdkkd");
        partner.setName("dldlkdk");
        Partner partner1 = new Partner();
        partner1.setAddress("dkkdkdkdkdk1");
        partner1.setBuildUpDate("ldldkdkdkkd1");
        partner1.setName("dldlkdk1");
        Partner partner2 = new Partner();
        partner2.setAddress("dkkdkdkdkdk2");
        partner2.setBuildUpDate("ldldkdkdkkd2");
        partner2.setName("dldlkdk2");
        partnerList.add(partner);
        partnerList.add(partner1);
        partnerList.add(partner2);
        contractInfoBean.setPartnerList(partnerList);
        InputStream inputStream = ParaseExcelTestMain.class.getResourceAsStream("/customerInfo.xls");
        OutputStream outputStream = new ByteArrayOutputStream();
        BuildExcelUtil.exportExcelByTemplate(inputStream,contractInfoBean,ContractInfoBean.class,outputStream,0);
    }
}
