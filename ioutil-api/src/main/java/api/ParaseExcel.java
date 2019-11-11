package api;

import annotation.ParaseExcelAnnotation;
import exception.ParaseExcelException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ParaseExcel<T> {
    /**
     * 以xls结尾的excel
     */
    private static String EXTENSION_XLS = "xls";

    private static String EXTENSION_XLSX = "xlsx";

    /**
     * 以路径的方式读取excel文件
     * 并将其解析成Javabean
     * 利用Java反射机制
     *
     * @param path
     */
    public List<T> readExcel(String path, Class<T> clazz) {
        if (null == path || "".equals(path)) {
            throw new ParaseExcelException("The file path is null");
        }
        File excFile = new File(path);
        if (!excFile.exists()) {
            throw new ParaseExcelException("the file does not exist：" + path);
        }
        if (!(path.endsWith(EXTENSION_XLS) || path.endsWith(EXTENSION_XLSX))) {
            throw new ParaseExcelException("this file is not excel");
        }
        return excelRowToBean(null, excFile, clazz);
    }

    /**
     * @param inputStream 为关闭流，需要在流创建的地方关闭流
     */
    public List<T> readExcel(InputStream inputStream, Class<T> clazz) {
        return excelRowToBean(inputStream, null, clazz);

    }

    /**
     * @param inputStream
     * @param file
     * @param clazz
     * @return
     */
    private List<T> excelRowToBean(InputStream inputStream, File file, Class<T> clazz) {
        if (null == clazz) {
            throw new ParaseExcelException("clazz is null ");
        }
        if (null == inputStream && null == file) {
            throw new ParaseExcelException("inputStream or file are null ");
        }
        Workbook workbook = null;


        List<T> resultList = new ArrayList<T>();
        try {
            if (null == inputStream) {
                workbook = WorkbookFactory.create(file);
            } else {
                workbook = WorkbookFactory.create(inputStream);
            }
            //得到Excel工作表对象
            Sheet sheetAt = workbook.getSheetAt(0);
            //循环读取表格数据
            for (Row row : sheetAt) {
                //首行（即表头）不读取
                if (row.getRowNum() == 0) {
                    continue;
                }
                int cellNum = row.getLastCellNum();
                //=======================================================
                //这是使用注解setter方法的方式,可以对数据进行特殊的处理但是循环次数变多
                T t = clazz.newInstance();
               /* Method[] methods = clazz.getMethods();
                for (Method method : methods) {
                    if (null == method.getAnnotation(ExcelCellParam.class)) {
                        continue;
                    }
                    int index = method.getAnnotation(ExcelCellParam.class).index();
                    if (index < cellNum) {
                        String areaValue = row.getCell(index).getStringCellValue();
                        method.invoke(t, areaValue);
                    }
                }*/
                //===========================================================
                //======================================================
                //这里使用的是注解field方法,循环次数少
                Field[] fields = clazz.getDeclaredFields();
                for(Field field:fields){
                    if(null == field.getAnnotation(ParaseExcelAnnotation.class)){
                        continue;
                    }
                    int index = field.getAnnotation(ParaseExcelAnnotation.class).index();
                    if (index < cellNum) {
                        String areaValue = row.getCell(index).getStringCellValue();
                        field.setAccessible(true);
                        field.set(t,areaValue);
                    }
                }
                //将循环一行的结果添加到返回的结果中
                resultList.add(t);
            }

        } catch (InvalidFormatException e) {
            e.printStackTrace();
            throw new ParaseExcelException("create a workbook object exception");
        } catch (IOException e) {
            e.printStackTrace();
            throw new ParaseExcelException("create a hssfWorkbook object exception");
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            throw new ParaseExcelException("create a bean object exception");
        } catch (InstantiationException e) {
            e.printStackTrace();
            throw new ParaseExcelException("create a bean object exception");
        } catch (Exception e){
            e.printStackTrace();
            throw new ParaseExcelException("Parase excel exception");
        }
            /*catch (InvocationTargetException e) {
            e.printStackTrace();
            throw new ParaseExcelException("reflect method invoke exception");
        }*/

        return resultList;

    }
}
