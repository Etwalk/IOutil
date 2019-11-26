package api;

import annotation.ExcelCellBeanAnnotation;
import annotation.ExcelCellListAnnotation;
import annotation.ExcelCellListBeginRowAnnotation;
import annotation.ExcelCellLocationAnnotation;
import exception.ParaseExcelException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
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
        InputStream inputStream = getFileByPath(path);
        return excelRowToBeans(inputStream, clazz, indexSheet);
    }

    /**
     * @param inputStream 为关闭流，需要在流创建的地方关闭流
     */
    public List readExcelBeans(InputStream inputStream, Class clazz, int indexSheet) {
        return excelRowToBeans(inputStream, clazz, indexSheet);

    }

    /**
     * 以路径的方式读取excel文件
     * 并将其解析成Javabean
     * 利用Java反射机制
     *
     * @param path
     */
    public Object readExcelBean(String path, Class clazz, int indexSheet) {
        InputStream inputStream = getFileByPath(path);
        return excelRowToBean(inputStream, clazz, indexSheet);
    }

    /**
     * @param inputStream 为关闭流，需要在流创建的地方关闭流
     */
    public Object readExcelBean(InputStream inputStream, Class clazz, int indexSheet) {
        return excelRowToBean(inputStream, clazz, indexSheet);

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
     * @param clazz
     * @return
     */
    private static Object excelRowToBean(InputStream inputStream, Class clazz, int sheetIndex) {
        if (null == clazz) {
            throw new ParaseExcelException("clazz is null ");
        }
        if (null == inputStream) {
            throw new ParaseExcelException("inputStream  is null ");
        }
        Workbook workbook = null;
        Object t = null;
        try {

            workbook = WorkbookFactory.create(inputStream);

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


    private static List excelRowToBeans(InputStream inputStream, Class clazz, int sheetIndex) {
        if (null == clazz) {
            throw new ParaseExcelException("clazz is null ");
        }
        if (null == inputStream) {
            throw new ParaseExcelException("inputStream is null ");
        }
        Workbook workbook = null;
        List resultList = new ArrayList();
        try {

            workbook = WorkbookFactory.create(inputStream);

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

    /**
     * 都转化成stream 就好了啊
     *
     * @param path
     * @return
     */
    private InputStream getFileByPath(String path) {
        if (null == path || "".equals(path)) {
            throw new ParaseExcelException("The file path is null");
        }
        return ParseExcel.class.getResourceAsStream("/customerInfo.xls");
    }

    private static void setObjectFieldStringValue(Cell cell,Object obj,Field field)throws IllegalAccessException {
        field.setAccessible(true);
        if(Cell.CELL_TYPE_STRING ==cell.getCellType()){
          field.set(obj,cell.getStringCellValue());
        }else if(Cell.CELL_TYPE_NUMERIC == cell.getCellType()){
            if(HSSFDateUtil.isCellDateFormatted(cell)){
                Date date = cell.getDateCellValue();
                field.set(obj,new SimpleDateFormat( "yyyy-MM-dd").format(date));
            }else{
                //这个地方怎么解决自动添加小数点的问题？
                double cellValue = cell.getNumericCellValue();
                field.set(obj,String.valueOf(cellValue));
            }
        }

    }
    //怎么样才能用一个方法来解决bean填入excel中和excel 读取到bean中即便有List

    public static Object parseExcelToObject(InputStream inputStream, Class clazz, int sheetIndex){
        if (null == clazz) {
            throw new ParaseExcelException("clazz is null ");
        }
        if (null == inputStream) {
            throw new ParaseExcelException("inputStream  is null ");
        }
        Workbook workbook = null;
        Object t = null;
        try {

            workbook = WorkbookFactory.create(inputStream);

            Sheet sheetAt = workbook.getSheetAt(sheetIndex);
            Field[] fields = clazz.getDeclaredFields();


            if(null != clazz.getDeclaredAnnotation(ExcelCellListBeginRowAnnotation.class)){
                //直接循环
                return valueCycle((ExcelCellListBeginRowAnnotation)clazz.getDeclaredAnnotation(ExcelCellListBeginRowAnnotation.class),sheetAt,clazz);
            }else{
                t = clazz.newInstance();
                for(Field field : fields){
                    if(null !=field.getAnnotation(ExcelCellBeanAnnotation.class)){
                        int rowNum = field.getAnnotation(ExcelCellBeanAnnotation.class).row();
                        int index = field.getAnnotation(ExcelCellBeanAnnotation.class).index();
                        Row row = sheetAt.getRow(rowNum);
                        Cell cell = row.getCell(index);
                        setObjectFieldStringValue(cell,t,field);
                    }else if(null != field.getAnnotation(ExcelCellListBeginRowAnnotation.class)&&"interface java.util.List".equals(field.getType().toString())){
                        int beginIndex = field.getAnnotation(ExcelCellListBeginRowAnnotation.class).beginRow();
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
                            List list = valueCycle(field.getAnnotation(ExcelCellListBeginRowAnnotation.class),sheetAt,genericClazz);
                            field.setAccessible(true);
                            field.set(t,list);
                        }


                    }

                }
                return t;

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

    }

    private static List valueCycle(ExcelCellListBeginRowAnnotation beginRowAnnotation,Sheet sheetAt,Class clazz){
        List resultList = new ArrayList();
        try {
            int beginRow = -1;
            if (beginRowAnnotation == null) {
                throw new ParaseExcelException(" ExcelCellListBeginRowAnnotation beginRow is error");
            } else {
                beginRow = beginRowAnnotation.beginRow();
            }
            Field[] fields = clazz.getDeclaredFields();
            for (Row row : sheetAt) {
                if (row.getRowNum() < beginRow) {
                    continue;
                }
                if (null == row.getCell(0) || "".equals(row.getCell(0).getStringCellValue())) {
                    break;
                }
                Object t = clazz.newInstance();
                for (Field field : fields) {
                    int index = field.getAnnotation(ExcelCellListAnnotation.class).index();
                    Cell cell = row.getCell(index);
                    setObjectFieldStringValue(cell,t,field);
                }
                //将循环一行的结果添加到返回的结果中
                resultList.add(t);
            }
        }catch(Exception e){
            e.printStackTrace();
            throw new ParaseExcelException("valueCycle exception");
        }
       return resultList;
    }


}
