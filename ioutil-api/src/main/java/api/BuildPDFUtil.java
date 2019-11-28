package api;

import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import exception.BuildPDFException;
import org.apache.poi.openxml4j.opc.OPCPackage;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;

/**
 * 定义了PDF中使用到的各种元素，其具体的整个pdf结构由此组成
 * 包括数据表格，各种可视化图。
 *
 * @author Sherlock.Wu
 * @date 2019/11/12
 */
public class BuildPDFUtil {
    //往模板中添数据
    public static void writeDataToPDFTemplate(String inPath, OutputStream outputStream, HashMap<String, String> templateMap) {
        try {
            PdfReader reader1 = new PdfReader(inPath);// 读取pdf模板
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            PdfStamper stamper = new PdfStamper(reader1, bos);
            AcroFields form = stamper.getAcroFields();//获取form域
            BaseFont bfChinese = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
            form.addSubstitutionFont(bfChinese);
            Iterator<String> iterator = form.getFields().keySet().iterator();
            int i = 0;
            while (iterator.hasNext()) {
                String name = iterator.next();
                String key = String.valueOf(i);//form域我设置名称为0-11,方便循环set值
                form.setField(name, templateMap.get(key));
                i++;
            }
            stamper.setFormFlattening(true);
            stamper.close();
            outputStream.write(bos.toByteArray());
            outputStream.flush();
            outputStream.close();
            bos.close();
        }catch (Exception e){
            e.printStackTrace();
            throw new BuildPDFException();
        }

    }

    //生成PDF并填写数据
    public static void createPDF() {

    }


    //往现存PDF中追加数据
    public static void writeDataToPDF() {

    }
}
