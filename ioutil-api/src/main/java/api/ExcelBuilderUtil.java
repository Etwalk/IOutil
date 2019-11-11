package api;

import annotation.ExcelBuilderAnnotation;
import exception.ExcelBuilderException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Iterator;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ExcelBuilderUtil<T> {
    public void exportExcel(String title, Collection<T> dateset, OutputStream out,Class clazz){
        try {
            //声明一个工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            //生成一个表格
            HSSFSheet hssfSheet = workbook.createSheet(title);
            //设置表格默认列宽度为18个字节
            hssfSheet.setDefaultColumnWidth(18);

            HSSFCellStyle style = getTitleStyle(workbook);

            HSSFCellStyle style1 = getBodyStyle(workbook);
            HSSFRow row = hssfSheet.createRow(0);

            Field[] fields = clazz.getDeclaredFields();

            int headCellIndex = 0;
            for(Field field:fields){
                if(null !=field.getAnnotation(ExcelBuilderAnnotation.class)){

                    HSSFCell cell = row.createCell(headCellIndex);
                    headCellIndex++;
                    cell.setCellStyle(style);
                    HSSFRichTextString text = new HSSFRichTextString(field.getAnnotation(ExcelBuilderAnnotation.class).value());
                    cell.setCellValue(text);
                }

            }
            Iterator<T> it = dateset.iterator();
            int index =0;
            while (it.hasNext()){
                index++;
                row = hssfSheet.createRow(index);

                T t = it.next();
                int i=0;
                for(Field field: fields){
                    if(null != field.getAnnotation(ExcelBuilderAnnotation.class)){
                        String fieldName = field.getName();
                        HSSFCell cell = row.createCell(i);
                        cell.setCellStyle(style1);
                        field.setAccessible(true);
                        Object obj = field.get(t);
                        String textValue = null;
                        if(null == obj){
                            textValue ="";
                        }else {
                            String classType = String.valueOf(obj.getClass());
                            if("class java.util.Date".equals(classType)){
                                String pattern ="yyyy-MM-dd HH:mm:ss";
                                SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                                textValue = sdf.format(obj);

                            }else {
                                textValue = String.valueOf(obj);
                            }
                        }
                        if(null != textValue  && textValue !=""){
                            HSSFRichTextString richTextString = new HSSFRichTextString(textValue);
                            cell.setCellValue(richTextString);
                        }
                        i++;
                    }

                }
            }

           workbook.write(out);
        }catch (IllegalAccessException e){
            e.printStackTrace();
            throw new ExcelBuilderException("set field accessible exception");
        }catch (Exception e){
            e.printStackTrace();
            throw new ExcelBuilderException("excel builder exception");
        }
    }

    public void exportExcelToFile(String title, Collection<T> dateset, String  filePath,Class clazz){
        try {
            File file = new File(filePath);
            OutputStream outputStream = new FileOutputStream(file);
            exportExcel(title,dateset,outputStream,clazz);
        }catch (IOException e){
            e.printStackTrace();
            throw new ExcelBuilderException("create fileoutputstream exception");
        }


    }









    private HSSFCellStyle getTitleStyle(HSSFWorkbook workbook){
        HSSFCellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.WHITE.index);
        font.setFontHeightInPoints(Short.parseShort("12"));
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontName("微软雅黑");
        style.setFont(font);
        return style;
    }

    private HSSFCellStyle getBodyStyle(HSSFWorkbook workbook){
        HSSFCellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontName("微软雅黑");
        style.setFont(font);
        return style;
    }
}
