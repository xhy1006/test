package com.example.sbsb.controller;

import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.CollectionUtils;

import javax.xml.bind.ValidationException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class excelread <T>{

 private Class claze;
 public excelread(Class claze){
     this.claze=claze;
 }

    /**
     * 基于注解读excel
     *
     * 文件地址   file
  *   从第几行开始读 rowInex
     *
     */

    public List<T> read(InputStream file,int rowIndex) throws Exception {

        List<T> list = new ArrayList<>();
        T entity =null;

            //带注解排列好的字段
        try {
            //获得数据vo类的注解字段
            List<Field> fieldList =getFieldList();
            Field field=null;
            Workbook workbook= WorkbookFactory.create(file);
            Sheet sheet=workbook.getSheetAt(0);
            // 表总行数
            int rowLength =sheet.getLastRowNum();
            for (int i=0;i< rowLength-2;i++){
                //第几行开始读
                Row row =sheet.getRow(i+rowIndex-1);
                int lastCellNum =row.getLastCellNum();
                // 得到数据vo的反射对象
                entity =(T)claze.newInstance();
                for (int j=0;j<lastCellNum-1;j++){
                    Cell cell =row.getCell(j+1);
                    field=fieldList.get(j);
                    if (cell==null){
                        // 字段为空  field.getName();
                        continue;
                    }
                    // 把对应的注解字段和cell内容对vo反射对象赋值；
                    field.set(entity,covertAttrType(field,cell));

                }
                list.add(entity);
            }


        } catch (Exception e) {

            throw new Exception("加载excel失败");
        }

        return list;

    }


    /**
     *
     * 获取带注解的字段
     *
     * @return
     */


    public List<Field> getFieldList() throws ValidationException {
        Field[] fields = this.claze.getDeclaredFields();
        //无序
        List<Field> fieldList =new ArrayList<>();
        //排序字段
        List<Field> fieldSortList= new LinkedList<>();
        int length =fields.length;
        int sort =0;
        Field field=null;

        //获取带注解字段
        for (int i=0;i<length;i++){

            field=fields[i];
            //判断元素是否有注解
            if(field.isAnnotationPresent(ExcelVo.class)){
                fieldList.add(field);
            }
        }

        if (CollectionUtils.isEmpty(fieldList)){

            throw  new ValidationException("未获取需要导入的字段");
        }
        length=fieldList.size();

        for (int i=1;i<=length;i++){

            //i==1 因为数据vo类属性序号从1开始的
            for (int j=0;j<length;j++){
                field=fieldList.get(j);
                //这一步步明白
                 ExcelVo excel = field.getAnnotation(ExcelVo.class);
                //让我们用反射时可以访问私有变量
                field.setAccessible(true);
                sort=excel.sort();
                if (sort==i){
                 fieldSortList.add(field);
                 continue;
                }

            }
        }

        return fieldSortList;
    }

    /**
     * 将 cell单元格格式转为 字段类型
     *
     */
    public Object covertAttrType(Field field, Cell cell) throws Exception {
        int type=cell.getCellType();
        if (type== Cell.CELL_TYPE_BLANK){
            return null;
        }
        ExcelVo excel = field.getAnnotation(ExcelVo.class);

        //字段类型
        String fieldtype =field.getType().getSimpleName();

        try {
            if ("String".equals(fieldtype)) {
                return getvalue(cell);
            } else if ("Date".equals(fieldtype)) {
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    Date cell1 = cell.getDateCellValue();
                    return cell;
                }
                return DateUtils.parseDate(getvalue(cell), new String[]{"yyyy-MM-dd"});
            } else if ("int".equals(fieldtype)) {
                return Integer.parseInt(getvalue(cell));
            } else if ("double".equals(fieldtype)) {
                return Double.parseDouble(getvalue(cell));
            }
        }catch (Exception e){
            if (e instanceof ParseException){
                // field.getName() 时间格式转换失败
            }else {
                // 格式转换失败
            }
        }

        throw new Exception("excel 单元格格式不支持");

    }

    /**
     *
     * 格式转string
     * @param cell
     * @return
     */

    public String getvalue(Cell cell){
        if (cell==null){
            return "";
        }
        switch (cell.getCellType()){
            case HSSFCell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString().trim();
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)){
                    Date dt =HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                    return DateFormatUtils.format(dt,"yyyy-MM-dd");
                }else {
                    //防止数值变为科学计数法
                    String strcell="";
                    Double num=cell.getNumericCellValue();
                    BigDecimal bd =new BigDecimal(num.toString());
                    if (bd!=null){
                        strcell=bd.toPlainString();
                    }
                    //去浮点数 自动加.0
                    if (strcell.endsWith(".0")){
                        strcell=strcell.substring(0,strcell.indexOf("."));
                    }
                    return strcell;
                }
            case HSSFCell.CELL_TYPE_BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case HSSFCell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            case HSSFCell.CELL_TYPE_BLANK:
                return "";
             default:
                 return "";
        }

    }
}
