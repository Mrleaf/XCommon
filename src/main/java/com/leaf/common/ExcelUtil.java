package com.leaf.common;


import com.leaf.model.ExcelInfo;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by leaf
 * 时间 2017/5/24 0024 14:30
 */
public class ExcelUtil {
    public static void main(String[] args) throws Exception{

    }

    /**
     * 判断从第几行读取数据
     * @param filePath  文件路径
     * @param start     开始行
     * @param validateList 验证字段
     * @param objClass  导入类
     * @param map   错误集合
     * @return
     */
    public static List<?> readExcel(String filePath, int start, List<ExcelInfo> validateList, Class<?> objClass,
                                    Map<String,Object> map) {
        Workbook book = null;
        List list = new ArrayList();
        try{
            book = getExcelWorkbook(filePath);
            if (book != null) {
                boolean flg = readHeader(book,0,validateList);
                if(flg){
                    map.put("error","上传文件不符合要求请下载模板重新上传");
                    map.put("isError",true);
                    return list;
                }
                list = readExcel(book,start,filePath,validateList,objClass,map);
            }
            return list;
        }catch (Exception e){
            e.printStackTrace();
            map.put("error","读取文件失败");
            map.put("isError",true);
            return list;
        }
    }

    /**
     * 校验表头
     * @param book
     * @param firstRowNum
     * @param validateList
     * @return
     * @throws Exception
     */
    private static boolean readHeader(Workbook book, int firstRowNum,List<ExcelInfo> validateList) throws Exception{
        Sheet sheet = getSheetByNum(book, 0);
        boolean isError = false;
        int maxCellNum = validateList.size();
        Row row = sheet.getRow(firstRowNum);
        if (row != null) {
            Cell cell = null;
            for (int j = 0; j < maxCellNum; j++) {
                cell = row.getCell(j);
                if (cell != null) {
                    ExcelInfo excelInfo = getValue(cell,validateList.get(j));
                    excelInfo.setError(false);
                    if(!excelInfo.getColumnsName().equals(excelInfo.getValue())){
                        isError = true;
                        break;
                    }

                }
            }

        }
        return  isError;
    }

    /**
     * 读取Excel表格
     * @param book  Excel表
     * @param firstRowNum   开始行
     * @param filePath  文件路径
     * @param validateList  验证字段
     * @param objClass  导入类
     * @param map   错误集合
     * @return
     */
    private static List<?> readExcel(Workbook book, int firstRowNum, String filePath,
                                     List<ExcelInfo> validateList, Class<?> objClass, Map<String,Object> map) throws Exception{
        Sheet sheet = getSheetByNum(book, 0);
        List<Object> listRow = new ArrayList();
        Object obj = null;
        Field field = null;
        boolean isEmpty = true,isCellError = false,isError = false;
        int count = sheet.getLastRowNum();
        int maxCellNum = validateList.size();
        Row row = null;
        for (int i = firstRowNum; i <= count; i++) {
            row = sheet.getRow(i);
            if (row != null) {
                int lastCellNum = row.getLastCellNum();
                Cell cell = null;
                obj = objClass.newInstance();
                isEmpty = true;
                isCellError = false;
                for (int j = 0; j < maxCellNum; j++) {
                    cell = row.getCell(j);
                    if (cell != null) {
                        ExcelInfo excelInfo = getValue(cell,validateList.get(j));
                        field = objClass.getDeclaredField(excelInfo.getColumnsCode());
                        field.setAccessible(true);
                        if(!ValidUtil.isNullOrEmptyStr(excelInfo.getValue())){
                            if(field.getType() == String.class&&"Number".equals(excelInfo.getType())){
                                excelInfo.setError(false);
                                field.set(obj,excelInfo.getValue());
                            }else if(field.getType() == Date.class&&"Date".equals(excelInfo.getType())){
                                if(!excelInfo.isError()){
                                    field.set(obj,DateUtil.toDate(excelInfo.getValue(),"yyyy-MM-dd HH:mm:ss"));
                                }
                            }else if(field.getType() == Integer.class){
                                if(excelInfo.isMoney()){
                                    int x = ValidUtil.isNullOrEmpty(excelInfo.getValue())?null:(int)(Double.valueOf(excelInfo.getValue())*100);
                                    field.set(obj,x);
                                }else{
                                    field.set(obj,ValidUtil.isNullOrEmpty(excelInfo.getValue())?null:Integer.valueOf(excelInfo.getValue()));
                                }
                            }else if(field.getType() ==  Long.class){
                                if(excelInfo.isMoney()){
                                    long x = ValidUtil.isNullOrEmpty(excelInfo.getValue())?null:(long)(Double.valueOf(excelInfo.getValue())*100);
                                    field.set(obj,x);
                                }else{
                                    field.set(obj,ValidUtil.isNullOrEmpty(excelInfo.getValue())?null:Integer.valueOf(excelInfo.getValue()));
                                }
                            }else if(field.getType() == Double.class){
                                field.set(obj,ValidUtil.isNullOrEmpty(excelInfo.getValue())?null:Double.valueOf(excelInfo.getValue()));
                            }else if(field.getType() == Short.class){
                                field.set(obj,ValidUtil.isNullOrEmpty(excelInfo.getValue())?null:Short.valueOf(excelInfo.getValue()));
                            }else{
                                field.set(obj,excelInfo.getValue());
                            }
                            isEmpty = false;
                        }
                        if(excelInfo.isError()){
                            isError = true;
                            isCellError = true;
                            CellStyle cellStyle = book.createCellStyle();
//                                cellStyle = cell.getCellStyle();
                            cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                            cell.setCellStyle(cellStyle);
                            cell.setCellValue(excelInfo.getValue());
                        }
                    }else {
                        cell = row.createCell(j);
                    }
                }
                if(!isEmpty)
                    listRow.add(obj);

            }
        }
        map.put("isError",isError);
        if(isError){
            FileOutputStream fileOut = new FileOutputStream(filePath);
            book.write(fileOut);
            fileOut.close();
        }

        return listRow;

    }

    /**
     * Excel模板
     *
     * @param tableHeader
     * @param savePath
     */
    public static boolean excelTemplet(List<ExcelInfo> tableHeader, String savePath) {
        try {
            SXSSFWorkbook wb = new SXSSFWorkbook();
            CreationHelper createHelper = wb.getCreationHelper();
            Sheet sheet1 = wb.createSheet("sheet1");
            Row row = sheet1.createRow((short) 0);
            Font font = wb.createFont();
            font.setFontName("宋体");
            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            font.setFontHeightInPoints((short) 13);
            CellStyle cs = wb.createCellStyle();
            cs.setFont(font);
            Cell cell = null;
            for (int i = 0; i < tableHeader.size(); i++) {
                ExcelInfo info = tableHeader.get(i);
                cell = row.createCell(i);
                cell.setCellValue(info.getColumnsName());
                cell.setCellStyle(cs);
            }
            File file = new File(savePath);
            if(!file.getParentFile().exists()) {
                //如果目标文件所在的目录不存在，则创建父目录
                if(!file.getParentFile().mkdirs()) {
                    return false;
                }
            }
            FileOutputStream fileOut = new FileOutputStream(savePath);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    /**
     * 导出数据
     * @param path      文件路径    例：C:\Users\Administrator\Desktop
     * @param fileName  文件名      例：导出测试
     * @param tableHeader   列名称信息
     * @param dataList  数据信息
     * @return  返回信息数据  path:文件路径   errorMsg：错误信息 isError：是否有错
     */
    public static Map<String,Object> exportExcel(String path,String fileName,List<ExcelInfo> tableHeader,List<?> dataList,Map<String,Map<String,String>> stateMap){
        return exportExcel(path,fileName,0,tableHeader,dataList,stateMap);
    }
    /**
     * 导出数据
     * @param path      文件路径    例：C:\Users\Administrator\Desktop
     * @param fileName  文件名      例：导出测试
     * @param tableHeader   列名称信息
     * @param dataList  数据信息
     * @return  返回信息数据  path:文件路径   errorMsg：错误信息 isError：是否有错
     */
    public static Map<String,Object> exportExcel(String path,String fileName,List<ExcelInfo> tableHeader,List<?> dataList){
        return exportExcel(path,fileName,0,tableHeader,dataList,null);
    }

    /**
     * 导出数据
     * @param path      文件路径    例：C:\Users\Administrator\Desktop
     * @param fileName  文件名      例：导出测试
     * @param tableHeader   列名称信息
     * @param dataList  数据信息
     * @param rowNum    从第几行写数据
     * @return  返回信息数据  path:文件路径   errorMsg：错误信息 isError：是否有错
     */
    public static Map<String,Object> exportExcel(String path,String fileName,int rowNum,List<ExcelInfo> tableHeader,
                                                 List<?> dataList,Map<String,Map<String,String>> stateMap){
        Map<String,Object> map = new HashMap<String,Object>();
        map.put("isError",false);
        try {
            String savePath = path+fileName+DateUtil.getDateStr("yyyyMMddHHmmSS")+".xls";
            SXSSFWorkbook wb = new SXSSFWorkbook();
            Sheet sheet = wb.createSheet("sheet1");

            Row row = sheet.createRow(rowNum);
            Font font = wb.createFont();
            font.setFontName("宋体");
            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            font.setFontHeightInPoints((short) 13);
            CellStyle cs = wb.createCellStyle();
            cs.setFont(font);
            Cell cell = null;
            for (int i = 0; i < tableHeader.size(); i++) {
                ExcelInfo info = tableHeader.get(i);
                sheet.setColumnWidth(i, 40*256);
//                sheet.autoSizeColumn(i, true);
                cell = row.createCell(i);
                cell.setCellValue(info.getColumnsName());
                cell.setCellStyle(cs);
            }
            Object obj = null;
            Map<String,String> tMap = null;
            for(int i=0;i<dataList.size();i++){
                obj = dataList.get(i);
                row = sheet.createRow(i+rowNum+1);
                for(int k=0;k<tableHeader.size();k++){
                    ExcelInfo excelInfo = tableHeader.get(k);
                    Object col = getFieldValueByName(excelInfo.getColumnsCode(),obj);
                    String value = "";
                    if(col instanceof Date){
                        value = ValidUtil.isNullOrEmpty(col)?"":DateUtil.toDateStr((Date)col,"yyyy/MM/dd HH:mm:ss");
                    }else if(col instanceof Integer){
                        if(!ValidUtil.isNullOrEmpty(col)&&excelInfo.isMoney()){
                            double x = (Integer)col/100.0;
                            value = String.valueOf(x);
                        }else{
                            value = ValidUtil.isNullOrEmpty(col)?"":col.toString();
                        }
                    }else if(col instanceof Long){
                        if(!ValidUtil.isNullOrEmpty(col)&&excelInfo.isMoney()){
                            double x = (Long)col/100.0;
                            value = String.valueOf(x);
                        }else{
                            value = ValidUtil.isNullOrEmpty(col)?"":col.toString();
                        }
                    }else{
                        value = ValidUtil.isNullOrEmpty(col)?"":col.toString();
                    }
//                    sheet.autoSizeColumn(k, true);
                    cell = row.createCell(k);
                    cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                    cell.setCellValue(value);
                }
            }
            File file = new File(savePath);
            if(!file.getParentFile().exists()) {
                if(!file.getParentFile().mkdirs()) {
                    map.put("errorMsg","创建失败");
                    map.put("isError",true);
                }
            }
            map.put("path",savePath);
            FileOutputStream fileOut = new FileOutputStream(savePath);
            wb.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
            map.put("isError",true);
            map.put("errorMsg","导出失败");
        }
        return map;
    }

    private static Object getFieldValueByName(String fieldName, Object o) {
        try {
            Class<?> objClass = o.getClass();
            Field[] fs = objClass.getDeclaredFields();
            Object val = null;
            for(int i = 0 ; i < fs.length; i++){
                Field f = fs[i];
                f.setAccessible(true); //设置些属性是可以访问的
                String name = f.getName();
                if(name.equals(fieldName)){
                    val = f.get(o);//得到此属性的值
                    break;
                }
            }
            return val;
        } catch (Exception e) {
            return null;
        }
    }


    /**
     * 根据索引返回 Sheet
     * @param book
     * @param number 从0开始  0表示第一个sheet
     * @return
     */
    private static Sheet getSheetByNum(Workbook book, int number) throws Exception{
        Sheet sheet = book.getSheetAt(number);
        return sheet;
    }

    /**
     * 获取文件
     *
     * @param filePath
     * @return
     * @throws IOException
     */
    public static Workbook getExcelWorkbook(String filePath) throws Exception {
        Workbook book = null;
        File file = null;
        FileInputStream fis = null;
        file = new File(filePath);
        if (!file.exists()) {
            throw new RuntimeException("没找到该文件");
        } else {
            fis = new FileInputStream(file);
            book = WorkbookFactory.create(fis);
        }

        if (fis != null) {
            fis.close();
        }
        return book;
    }


    /**
     * 返回 value
     *
     * @param cell
     * @return
     */
    private static ExcelInfo getValue(Cell cell,ExcelInfo excelInfo) throws Exception {
        String value = "";
        int cellType = cell.getCellType();
        String type_cn = null;
        switch (cellType) {
            case Cell.CELL_TYPE_NUMERIC:
                //判断数值类型
                if(HSSFDateUtil.isCellDateFormatted(cell)||isReserved(cell.getCellStyle().getDataFormat())
                        ||isDateFormat(cell.getCellStyle().getDataFormatString())){
                    type_cn = "Date";
                    Date date = cell.getDateCellValue();
                    value = DateUtil.toDateStr(date,"yyyy-MM-dd HH:mm:ss");
                }else{
                    type_cn = "Number";
                    DecimalFormat dec = new DecimalFormat("0");
                    value = dec.format(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_STRING:
                //字符串型
                type_cn = "String";
                value = cell.getStringCellValue();
                String format = isDateGetFormat(value);
                if(!ValidUtil.isNullOrEmptyStr(format)){
                    type_cn = "Date";
                    value = DateUtil.toDateStr(DateUtil.toDate(value,format),"yyyy-MM-dd HH:mm:ss");
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                //公式型
                type_cn = "Formula";
                value = cell.getCellFormula();
                break;
            case Cell.CELL_TYPE_BLANK:
                //空
                type_cn = "Blank";
                value = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                //布尔型
                type_cn = "Boolean";
                boolean tempValue = cell.getBooleanCellValue();
                value = String.valueOf(tempValue);
                break;
            case Cell.CELL_TYPE_ERROR:
                //错误
                type_cn = "Error";
                byte b = cell.getErrorCellValue();
                value = String.valueOf(b);
                break;
            default:
                type_cn = "String";
                value = cell.getStringCellValue();
                break;
        }
        value = value.trim();
        excelInfo.setValue(value);
        if(ValidUtil.isNullOrEmptyStr(value)){
            excelInfo.setError(false);
        }else if("String".equals(type_cn)&&excelInfo.getType().equals(type_cn)
                &&value.length()<=excelInfo.getLength()){
            excelInfo.setError(false);
        }else if("Number".equals(type_cn)&&excelInfo.getType().equals(type_cn)){
            excelInfo.setError(false);
        }else if("Date".equals(type_cn)&&excelInfo.getType().equals(type_cn)){
            excelInfo.setError(false);
        }else if("Number".equals(type_cn)&&"String".equals(excelInfo.getType())&&!excelInfo.getType().equals(type_cn)){
            excelInfo.setError(false);
        }else if("String".equals(type_cn)&&"Number".equals(excelInfo.getType())&&!excelInfo.getType().equals(type_cn)){
            excelInfo.setError(false);
        }else if("Date".equals(type_cn)&&"String".equals(excelInfo.getType())&&!excelInfo.getType().equals(type_cn)){
            excelInfo.setError(false);
        }else{
            excelInfo.setError(true);
        }
        return excelInfo;
    }

    /**
     * 是否是日期格式保留字段
     * @param reserv
     * @return
     */
    private static boolean isReserved(short reserv){
        if(reserv>=27&&reserv<=31){
            return true;
        }
        return false;
    }
    /**
     * 判断是否是中文日期格式
     * @param isNotDate
     * @return boolean
     */
    private static boolean isDateFormat(String isNotDate){
        if(ValidUtil.isNullOrEmptyStr(isNotDate)){
            return false;
        }
        if(isNotDate.contains("年")||isNotDate.contains("月")||isNotDate.contains("日")||
                isNotDate.contains("时")||isNotDate.contains("分")||isNotDate.contains("秒")||
                isNotDate.toUpperCase().contains("AM")||isNotDate.toUpperCase().contains("PM")) {
            return true;
        }
        return false;
    }

    /**
     * 判断是否是日期并返回时间格式
     * @param date
     * @return
     */
    public static String isDateGetFormat(String date) {
        if(ValidUtil.isNullOrEmptyStr(date)){
            return "";
        }
        Map<String,SimpleDateFormat> formatMap = new LinkedHashMap<String,SimpleDateFormat>();
        if(date.contains("-")){
            formatMap.put("yyyy-MM-dd HH:mm:ss",new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"));
            formatMap.put("yyyy-MM-dd HH:mm",new SimpleDateFormat("yyyy-MM-dd HH:mm"));
            formatMap.put("yyyy-MM-dd HH",new SimpleDateFormat("yyyy-MM-dd HH"));
            formatMap.put("yyyy-MM-dd",new SimpleDateFormat("yyyy-MM-dd"));
            formatMap.put("yyyy-MM",new SimpleDateFormat("yyyy-MM"));
        }else if(date.contains("/")){
            formatMap.put("yyyy/MM/dd HH:mm:ss",new SimpleDateFormat("yyyy/MM/dd HH:mm:ss"));
            formatMap.put("yyyy/MM/dd HH:mm",new SimpleDateFormat("yyyy/MM/dd HH:mm"));
            formatMap.put("yyyy/MM/dd HH",new SimpleDateFormat("yyyy/MM/dd HH"));
            formatMap.put("yyyy/MM/dd",new SimpleDateFormat("yyyy/MM/dd"));
            formatMap.put("yyyy/MM",new SimpleDateFormat("yyyy/MM"));
        }else if(date.contains("年")){
            formatMap.put("yyyy年MM月dd日 HH时mm分ss秒",new SimpleDateFormat("yyyy年MM月dd日 HH时mm分ss秒"));
            formatMap.put("yyyy年MM月dd日 HH时mm分",new SimpleDateFormat("yyyy年MM月dd日 HH时mm分"));
            formatMap.put("yyyy年MM月dd日 HH时",new SimpleDateFormat("yyyy年MM月dd日 HH时"));
            formatMap.put("yyyy年MM月dd日",new SimpleDateFormat("yyyy年MM月dd日"));
            formatMap.put("yyyy年MM月",new SimpleDateFormat("yyyy年MM月"));
        }
        SimpleDateFormat format;
        String formatStr = "";
        for(Map.Entry<String,SimpleDateFormat> entry:formatMap.entrySet()){
            try {
                format = entry.getValue();
                format.parse(date);
                formatStr = entry.getKey();
                break;
            } catch (ParseException pe) {
                formatStr = "";
            }
        }
        return formatStr;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row 行下标
     * @param column 列下标
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取合并单元格的值
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet ,int row , int column){
        int sheetMergeCount = sheet.getNumMergedRegions();
        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell) ;
                }
            }
        }
        return null ;
    }

    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell){
        if(cell == null) return "";
        DecimalFormat df = new DecimalFormat("0");
        if(cell.getCellType() == Cell.CELL_TYPE_STRING){
            return cell.getStringCellValue();
        }else if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN){
            return String.valueOf(cell.getBooleanCellValue());
        }else if(cell.getCellType() == Cell.CELL_TYPE_FORMULA){
            return cell.getCellFormula() ;
        }else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
            return df.format(cell.getNumericCellValue());
        }else{
            return "";
        }
    }
}
