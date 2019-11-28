package api;

import annotation.BuildExcelAnnotation;
import annotation.ExcelCellBeanAnnotation;
import annotation.ExcelCellListAnnotation;
import annotation.ExcelCellListBeginRowAnnotation;
import exception.BuildExcelException;
import exception.ParaseExcelException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class BuildExcelUtil<T> {
    public void exportExcel(String title, Collection<T> dateset, OutputStream out, Class clazz) {
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
            for (Field field : fields) {
                if (null != field.getAnnotation(BuildExcelAnnotation.class)) {

                    HSSFCell cell = row.createCell(headCellIndex);
                    headCellIndex++;
                    cell.setCellStyle(style);
                    HSSFRichTextString text = new HSSFRichTextString(field.getAnnotation(BuildExcelAnnotation.class).value());
                    cell.setCellValue(text);
                }

            }
            Iterator<T> it = dateset.iterator();
            int index = 0;
            while (it.hasNext()) {
                index++;
                row = hssfSheet.createRow(index);

                T t = it.next();
                int i = 0;
                for (Field field : fields) {
                    if (null != field.getAnnotation(BuildExcelAnnotation.class)) {
                        String fieldName = field.getName();
                        HSSFCell cell = row.createCell(i);
                        cell.setCellStyle(style1);
                        field.setAccessible(true);
                        Object obj = field.get(t);
                        String textValue = null;
                        if (null == obj) {
                            textValue = "";
                        } else {
                            String classType = String.valueOf(obj.getClass());
                            if ("class java.util.Date".equals(classType)) {
                                String pattern = "yyyy-MM-dd HH:mm:ss";
                                SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                                textValue = sdf.format(obj);

                            } else {
                                textValue = String.valueOf(obj);
                            }
                        }
                        if (null != textValue && textValue != "") {
                            HSSFRichTextString richTextString = new HSSFRichTextString(textValue);
                            cell.setCellValue(richTextString);
                        }
                        i++;
                    }

                }
            }

            workbook.write(out);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            throw new BuildExcelException("set field accessible exception");
        } catch (Exception e) {
            e.printStackTrace();
            throw new BuildExcelException("excel builder exception");
        }
    }

    public void exportExcelToFile(String title, Collection<T> dateset, String filePath, Class clazz) {
        try {
            File file = new File(filePath);
            OutputStream outputStream = new FileOutputStream(file);
            exportExcel(title, dateset, outputStream, clazz);
        } catch (IOException e) {
            e.printStackTrace();
            throw new BuildExcelException("create fileoutputstream exception");
        }


    }


    private HSSFCellStyle getTitleStyle(HSSFWorkbook workbook) {
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

    private HSSFCellStyle getBodyStyle(HSSFWorkbook workbook) {
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

    /**
     * 根据路径获取到excel模板,然后写入数据
     * @param path
     * @param dateset
     * @param out
     * @param clazz
     * @param sheetIndex
     */
    public void exportExcelByTemplate(String path, Object dateset, OutputStream out, Class clazz, int sheetIndex) {
        if (null == path || "".equals(path)) {
            throw new ParaseExcelException("The file path is null");
        }
        InputStream inputStream = ParseExcel.class.getResourceAsStream("path");
        exportExcelByTemplate(inputStream, dateset, clazz, out, sheetIndex);
        InputStreamUtil.close(inputStream);
    }

    /**
     * 根据inputStream 获取到的excel模板然后写入数据
     * @param inputStream
     * @param dateset
     * @param clazz
     * @param out
     * @param sheetIndex
     */
    public static void exportExcelByTemplate(InputStream inputStream, Object dateset, Class clazz, OutputStream out, int sheetIndex) {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);

            Sheet sheetAt = workbook.getSheetAt(sheetIndex);
            System.out.println(workbook);
            if ("java.util.ArrayList".equals(dateset.getClass().getName())) {
                //这个地方直接循环
                if (null != clazz.getDeclaredAnnotation(ExcelCellListBeginRowAnnotation.class)) {
                    ExcelCellListBeginRowAnnotation beginRowAnnotation = (ExcelCellListBeginRowAnnotation) clazz.getDeclaredAnnotation(ExcelCellListBeginRowAnnotation.class);
                    int beginRow = beginRowAnnotation.beginRow();
                    if (beginRow < 0) {
                        throw new BuildExcelException("beginRow can not less zero");
                    }
                    writeListToExcel((List)dateset,beginRow,sheetAt,clazz);
                    //如果开始的不为行大于等于0则不加标题,即便有BuildExcelAnnotation也不解析使用。
                }
            } else if (dateset.getClass().getName().equals(clazz.getName())) {
                Field[] fields = clazz.getDeclaredFields();
                for (Field field : fields) {
                    if (null != field.getAnnotation(ExcelCellBeanAnnotation.class) && !("interface java.util.List".equals(field.getType().toString()))) {
                        int rowNum = field.getAnnotation(ExcelCellBeanAnnotation.class).row();
                        int index = field.getAnnotation(ExcelCellBeanAnnotation.class).index();
                        Row row = sheetAt.getRow(rowNum);
                        Cell cell = row.getCell(index);
                        field.setAccessible(true);
                        field.get(dateset);
                        System.out.println("field:" + field.get(dateset));
                        cell.setCellValue(field.get(dateset).toString());
                    } else if (null != field.getAnnotation(ExcelCellListBeginRowAnnotation.class) && "interface java.util.List".equals(field.getType().toString())) {
                        //循环遍历插入值
                        int beginIndex = field.getAnnotation(ExcelCellListBeginRowAnnotation.class).beginRow();
                        if (beginIndex < 0) {
                            throw new ParaseExcelException("list annotation beginIndex not init");
                        }
                        Type genericType = field.getGenericType();
                        if (genericType == null) {
                            throw new ParaseExcelException("please make sure generics of list");
                        }
                        // 如果是泛型参数的类型
                        if (genericType instanceof ParameterizedType) {
                            ParameterizedType pt = (ParameterizedType) genericType;
                            //得到泛型里的class类型对象
                            Class<?> genericClazz = (Class<?>) pt.getActualTypeArguments()[0];
                            field.setAccessible(true);
                            System.out.println(field.get(dateset));
                            List list = (List) field.get(dateset);
                            writeListToExcel(list,beginIndex,sheetAt,genericClazz);
                        }
                    }
                }
            }

            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
            throw new BuildExcelException("IOException");
        }

    }

    /**
     * 把list写进excel中
     * @param list
     * @param beginRow
     * @param sheet
     * @param clazz
     */
    private static void writeListToExcel(List list,int beginRow,Sheet sheet,Class clazz){
        if(null == list){
            throw new BuildExcelException("list is null");
        }
        try{
            Field[] fields = clazz.getDeclaredFields();
            for(int i = 0;i<list.size();i++){
                Object object =list.get(i);
                Row row =sheet.getRow(beginRow+i);
                for(Field field : fields){
                    if(null !=field.getAnnotation(ExcelCellListAnnotation.class)&& !("interface java.util.List".equals(field.getType().toString()))){
                        int index = field.getAnnotation(ExcelCellListAnnotation.class).index();
                        if(index>=0){
                            Cell cell = row.getCell(index);
                            field.setAccessible(true);
                            field.get(object);
                            System.out.println(field.get(object));
                            if(field.get(object)!=null){
                                cell.setCellValue(field.get(object).toString());
                            }

                        }
                    }
                }
            }
        }catch(Exception e){
            e.printStackTrace();
            throw new BuildExcelException();
        }


    }
    //这个默认加在excel的后面
    public static void addDataToExcelEnd(String path){
        if (null == path || "".equals(path)) {
            throw new ParaseExcelException("The file path is null");
        }
        InputStream inputStream = ParseExcel.class.getResourceAsStream("path");
        addDataToExcelEnd(inputStream);
        InputStreamUtil.close(inputStream);
    }
    public static void addDataToExcelEnd(InputStream inputStream){
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);

            Sheet sheetAt = workbook.getSheetAt(0);
        }catch (Exception e){
            e.printStackTrace();
            throw new BuildExcelException("addDataToExcelEnd exception");
        }
    }

}
