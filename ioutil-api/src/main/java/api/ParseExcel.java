package api;

import annotation.ExcelCellBeanAnnotation;
import annotation.ExcelCellListAnnotation;
import annotation.ExcelCellListBeginRowAnnotation;
import annotation.ExcelCellLocationAnnotation;
import exception.ParaseExcelException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Sherlock.Wu
 * @date 2019/11/11
 */
public class ParseExcel<T> {
    /**
     * 以xls结尾的excel
     */
    private static String EXTENSION_XLS = "xls";

    private static String EXTENSION_XLSX = "xlsx";

    private static int INIT_ERROR_CODE = -1;

    /**
     * 以路径的方式读取excel文件
     * 并将其解析成Javabean
     * 利用Java反射机制
     *
     * @param path
     */
    public List readExcelBeans(String path, Class clazz, int indexSheet) {
        File excFile = getFileByPath(path);
        return excelRowToBeans(null, excFile, clazz, indexSheet);
    }

    /**
     * @param inputStream 为关闭流，需要在流创建的地方关闭流
     */
    public List readExcelBeans(InputStream inputStream, Class clazz, int indexSheet) {
        return excelRowToBeans(inputStream, null, clazz, indexSheet);

    }

    /**
     * 以路径的方式读取excel文件
     * 并将其解析成Javabean
     * 利用Java反射机制
     *
     * @param path
     */
    public Object readExcelBean(String path, Class clazz, int indexSheet) {
        File excFile = getFileByPath(path);
        return excelRowToBean(null, excFile, clazz, indexSheet);
    }

    /**
     * @param inputStream 为关闭流，需要在流创建的地方关闭流
     */
    public Object readExcelBean(InputStream inputStream, Class clazz, int indexSheet) {
        return excelRowToBean(inputStream, null, clazz, indexSheet);

    }

    /**
     * @param inputStream
     * @param file
     * @param clazz
     * @return
     */
   /* private static List excelRowToBean(InputStream inputStream, File file, Class clazz) {
        if (null == clazz) {
            throw new ParaseExcelException("clazz is null ");
        }
        if (null == inputStream && null == file) {
            throw new ParaseExcelException("inputStream or file are null ");
        }
        Workbook workbook = null;


        List resultList = new ArrayList();
        try {
            if (null == inputStream) {
                workbook = WorkbookFactory.create(file);
            } else {
                workbook = WorkbookFactory.create(inputStream);
            }
            //得到Excel工作表对象
            int sheetNum = workbook.getActiveSheetIndex();
            System.out.println(sheetNum);
            Sheet sheetAt = workbook.getSheetAt(0);
            Field[] fields = clazz.getDeclaredFields();
            //以下对class的结构类型进行结构性的判断

            //循环读取表格数据
//            for (Row row : sheetAt) {
//                //首行（即表头）不读取
//                if (row.getRowNum() == 0) {
//                    continue;
//                }
//                int cellNum = row.getLastCellNum();
            //=======================================================
            //这是使用注解setter方法的方式,可以对数据进行特殊的处理但是循环次数变多
            Object t = clazz.newInstance();
               *//* Method[] methods = clazz.getMethods();
                for (Method method : methods) {
                    if (null == method.getAnnotation(ExcelCellParam.class)) {
                        continue;
                    }
                    int index = method.getAnnotation(ExcelCellParam.class).index();
                    if (index < cellNum) {
                        String areaValue = row.getCell(index).getStringCellValue();
                        method.invoke(t, areaValue);
                    }
                }*//*
            //===========================================================
            //======================================================
            //这里使用的是注解field方法,循环次数少

            for (Field field : fields) {
                if (null == field.getAnnotation(ExcelCellLocationAnnotation.class)) {
                    continue;
                }
                //确定是否是list
                if ("interface java.util.List".equals(field.getType().toString())) {
                    int beginIndex = field.getAnnotation(ExcelCellLocationAnnotation.class).beginRow();
                    Type genericType = field.getGenericType();
                    if (genericType == null) {
                        throw new ParaseExcelException("please make sure generics of list");
                    }
                    // 如果是泛型参数的类型
                    if (genericType instanceof ParameterizedType) {
                        ParameterizedType pt = (ParameterizedType) genericType;
                        //得到泛型里的class类型对象
                        Class<?> genericClazz = (Class<?>) pt.getActualTypeArguments()[0];
                        Object obj = genericClazz.newInstance();

                        System.out.println(obj);
                    }

                } else {
                    int row = field.getAnnotation(ExcelCellLocationAnnotation.class).row();
                    int index = field.getAnnotation(ExcelCellLocationAnnotation.class).index();
                }

//                    Row row1 = sheetAt.getRow(row);
//                    String areaValue = row1.getCell(index).getStringCellValue();
//                    field.setAccessible(true);
//                    field.set(t,areaValue);
//                    if (index < cellNum) {
//                        String areaValue = row.getCell(index).getStringCellValue();
//                        field.setAccessible(true);
//                        field.set(t,areaValue);
//                    }
            }
            //将循环一行的结果添加到返回的结果中
            resultList.add(t);
//            }

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
        } catch (Exception e) {
            e.printStackTrace();
            throw new ParaseExcelException("Parse excel exception");
        }
            *//*catch (InvocationTargetException e) {
            e.printStackTrace();
            throw new ParaseExcelException("reflect method invoke exception");
        }*//*

        return resultList;

    }*/

    /**
     * @param inputStream
     * @param file
     * @param clazz
     * @return
     */
    private static Object excelRowToBean(InputStream inputStream, File file, Class clazz, int sheetIndex) {
        if (null == clazz) {
            throw new ParaseExcelException("clazz is null ");
        }
        if (null == inputStream && null == file) {
            throw new ParaseExcelException("inputStream or file are null ");
        }
        Workbook workbook = null;
        Object t = null;
        try {
            if (null == inputStream) {
                workbook = WorkbookFactory.create(file);
            } else {
                workbook = WorkbookFactory.create(inputStream);
            }
            Sheet sheetAt = workbook.getSheetAt(sheetIndex);
            Field[] fields = clazz.getDeclaredFields();

            t = clazz.newInstance();

            for (Field field : fields) {
                if (null == field.getAnnotation(ExcelCellBeanAnnotation.class)) {
                    continue;
                }
                int rowNum = field.getAnnotation(ExcelCellBeanAnnotation.class).row();
                int index = field.getAnnotation(ExcelCellBeanAnnotation.class).index();
                Row row = sheetAt.getRow(rowNum);
                String areaValue = row.getCell(index).getStringCellValue();
                field.setAccessible(true);
                field.set(t, areaValue);
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
        } catch (Exception e) {
            e.printStackTrace();
            throw new ParaseExcelException("Parse excel exception");
        }

        return t;

    }


    private static List excelRowToBeans(InputStream inputStream, File file, Class clazz, int sheetIndex) {
        if (null == clazz) {
            throw new ParaseExcelException("clazz is null ");
        }
        if (null == inputStream && null == file) {
            throw new ParaseExcelException("inputStream or file are null ");
        }
        Workbook workbook = null;
        List resultList = new ArrayList();
        try {
            if (null == inputStream) {
                workbook = WorkbookFactory.create(file);
            } else {
                workbook = WorkbookFactory.create(inputStream);
            }
            Sheet sheetAt = workbook.getSheetAt(sheetIndex);
            int beginRow = -1;
            Field[] fields = clazz.getDeclaredFields();
            ExcelCellListBeginRowAnnotation beginRowAnnotation = (ExcelCellListBeginRowAnnotation) clazz.getAnnotation(ExcelCellListBeginRowAnnotation.class);
            if (beginRowAnnotation == null) {
                throw new ParaseExcelException(" ExcelCellListBeginRowAnnotation beginRow is error");
            } else {
                beginRow = beginRowAnnotation.beginRow();
            }


            for (Row row : sheetAt) {
                if (row.getRowNum() < beginRow) {
                    continue;
                }
                if (null == row.getCell(0) || "".equals(row.getCell(0).getStringCellValue())) {
                    break;
                }
                System.out.println(row.getRowNum());
                System.out.println(row.getCell(0));
                Object t = clazz.newInstance();
                for (Field field : fields) {
                    int index = field.getAnnotation(ExcelCellListAnnotation.class).index();
                    String areaValue = row.getCell(index).getStringCellValue();
                    field.setAccessible(true);
                    field.set(t, areaValue);
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
        } catch (Exception e) {
            e.printStackTrace();
            throw new ParaseExcelException("Parse excel exception");
        }

        return resultList;

    }




    /*private boolean objectStructureJudgment(Class clazz) {
        try {
            Field[] fields = clazz.getDeclaredFields();
            //以下对class的结构类型进行结构性的判断
            int dealMethod = 0;
            int beginRowValue = -1;
            for (Field field : fields) {

                if (null == field.getAnnotation(ExcelCellLocationAnnotation.class)) {
                    continue;
                }
                ExcelCellLocationAnnotation fieldAnnotation =field.getAnnotation(ExcelCellLocationAnnotation.class);
                int beginRow = fieldAnnotation.beginRow();
                int row = fieldAnnotation.row();
                int index = fieldAnnotation.index();
                if(INIT_ERROR_CODE == beginRow && INIT_ERROR_CODE == row && INIT_ERROR_CODE ==index){
                    throw  new ParaseExcelException("Error annotation value");
                }else if(INIT_ERROR_CODE == row && INIT_ERROR_CODE !=beginRow && INIT_ERROR_CODE != index){
                    //单纯的循环
                    if (dealMethod == 0 || dealMethod ==1){
                        dealMethod =1;
                    }else{
                        throw new ParaseExcelException("please make sure bean struct true");
                    }
                    dealMethod = 1;
                    if(beginRowValue == -1){
                        beginRowValue =beginRow;
                    }else{
                        if(beginRowValue !==beginRow){
                            throw new ParaseExcelException("beginRows not equal");
                        }
                    }
                }else if(INIT_ERROR_CODE != row &&INIT_ERROR_CODE ==beginRow && INIT_ERROR_CODE !=index){
                    if(dealMethod == 0 ||dealMethod == 2){
                        dealMethod =2;
                    }else{
                        throw new ParaseExcelException("please make sure bean struct true");
                    }
                }else if()

                if ("interface java.util.List".equals(field.getType().toString())) {
                    hasList = true;
                    int beginIndex = field.getAnnotation(ExcelCellLocationAnnotation.class).beginRow();
                    if(INIT_ERROR_CODE == beginIndex){
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
                        Object obj = genericClazz.newInstance();

                        System.out.println(obj);
                        boolean result = objectStructureJudgment(genericClazz);
                        if (!result) {
                           throw new ParaseExcelException("bean object struct exception");
                        }
                    }

                } else if ("class java.lang.String".equals(field.getType().toString())) {

                } else {
                    throw new ParaseExcelException(clazz.getName() + "Here are the types that do not meet the requirements");
                }
            }
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            throw new ParaseExcelException("bean object struct exception");
        }

    }*/

    private File getFileByPath(String path) {
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
        return excFile;
    }

    public static byte[] getBytes(InputStream inputStream) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len;
        byte[] dataBytes;
        while ((len = inputStream.read(buffer)) != -1) {
            baos.write(buffer, 0, len);
        }
        baos.flush();
        dataBytes = baos.toByteArray();
        return dataBytes;
    }

    public static InputStream getNewStream(InputStream inputStream) throws IOException {
        byte[] dataBytes = getBytes(inputStream);
        return new ByteArrayInputStream(dataBytes);
    }
}
