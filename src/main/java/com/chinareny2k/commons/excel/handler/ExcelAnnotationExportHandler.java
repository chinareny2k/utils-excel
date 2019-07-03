package com.chinareny2k.commons.excel.handler;


import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang.ArrayUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.chinareny2k.commons.excel.ExcelExportHandler;
import com.chinareny2k.commons.excel.ExcelExportMethod;
import com.chinareny2k.commons.excel.annotation.Excel;
import com.chinareny2k.commons.excel.annotation.ExportField;

public class ExcelAnnotationExportHandler<item> implements ExcelExportHandler<item> {

  private final Class<item> entityClass;
  
  private static Logger logger = Logger.getLogger(ExcelAnnotationExportHandler.class);
  
  public static String EXCEL_ANNOTATION_MAPPING_DEFAULT = "default";
  public static String EXCEL_ANNOTATION_MAPPING_SPLIT_PROPERITY = "\\|\\|";
  public static String EXCEL_ANNOTATION_MAPPING_SPLIT_KEYVALUE = "\\$\\$";

  ExcelAnnotationExportHandler(Class<item> entityClass) {
    this.entityClass = entityClass;
  }



  public void export(List<item> list, OutputStream out) throws Exception {
    List<ExcelExportMethod> methods = getExcelExportMethods();
    if (CollectionUtils.isEmpty(methods)) {
      logger.warn("Excel导出异常 : 未设置Excel导出功能不能进行导出");
      throw new Exception("未设置Excel导出功能不能进行导出");
    }

    Excel excel = entityClass.getAnnotation(Excel.class);

    HSSFWorkbook workbook = new HSSFWorkbook();
    Sheet sheet = workbook.createSheet(null != excel ? excel.sheetName() : "Sheet1");
    writeHead(sheet, methods);
    writeItems(sheet, methods, list);
    try {
      workbook.write(out);
    } catch (IOException e) {
      e.printStackTrace();
      logger.error("Excel导出异常",e);
    }finally{
    	if(workbook!=null){
    		workbook.close();
    	}
    }
  }

  /**
   * 写excel的内容
   * 
   * @param sheet
   * @param methods
   * @param list
   */
  private void writeItems(Sheet sheet, List<ExcelExportMethod> methods, List<item> list) {
    int n = 1;
    for (item excelItem : list) {
      Row row = sheet.createRow(n);
      for (int i = 0; i < methods.size(); i++) {
        Cell cell = row.createCell(i);
        cell.setCellValue(getCellValue(methods.get(i),excelItem));
      }
      n++;
    }
  }

  /**
   * 设置Excel单元格的值
   * 
   * @param excelExportMethod
   * @param cell
   * @param excelItem
   */
  private String getCellValue(ExcelExportMethod excelExportMethod, item excelItem) {
    // 获取值的类型
    Method readMethod = excelExportMethod.getReadMethod();
    try {
      Object value = readMethod.invoke(excelItem, new Object[0]);
      String mapping = excelExportMethod.getMapping();
      return getValue(value, excelExportMethod.getFormat(), mapping);
    } catch (Exception e) {
      e.printStackTrace();
      logger.error(e);
      return "";
    }

  }

  private String getValue(Object value, String format, String mapping) {
    Map<String, String> properities = getMapping(mapping);
    if (MapUtils.isEmpty(properities)) {
      return getValue(value, format);
    } else {
      if (properities.containsKey(value.toString())) {
        return properities.get(value.toString());
      } else if (properities.containsKey(EXCEL_ANNOTATION_MAPPING_DEFAULT)) {
        return properities.get(EXCEL_ANNOTATION_MAPPING_DEFAULT);
      } else {
        return getValue(value, format);
      }
    }
  }

  private Map<String, String> getMapping(String mapping) {
    Map<String, String> map = new HashMap<String, String>();
    if (StringUtils.isBlank(mapping)) {
      return map;
    }
    String[] mappings = mapping.split(EXCEL_ANNOTATION_MAPPING_SPLIT_PROPERITY);
    for (String str : mappings) {
      String[] property = str.split(EXCEL_ANNOTATION_MAPPING_SPLIT_KEYVALUE);
      if (ArrayUtils.isNotEmpty(property) && 2 == property.length) {
        map.put(property[0], property[1]);
      }
    }
    return map;
  }

  private String getValue(Object value, String format) {
    String textValue = "";
    if (value == null)
      return textValue;

    if (value instanceof Boolean) {
      boolean bValue = (Boolean) value;
      textValue = EXCEL_BOOLEAN_MAPPING_YES;
      if (!bValue) {
        textValue = EXCEL_BOOLEAN_MAPPING_NO;
      }
    } else if (value instanceof Date) {
      Date date = (Date) value;
      SimpleDateFormat sdf = new SimpleDateFormat(format);
      textValue = sdf.format(date);
    } else
      textValue = value.toString();

    return textValue;

  }

  /**
   * 写excel的标题头
   * 
   * @param sheet
   * @param methods
   */
  private void writeHead(Sheet sheet, List<ExcelExportMethod> methods) {
    Row row = sheet.createRow(0);
    for (int i = 0; i < methods.size(); i++) {
      sheet.setColumnWidth(i, (int)(methods.get(i).getColumnWidth() * EXCEL_COLUMNWIDTH_PIXELCONVERSION));
      row.createCell(i).setCellValue(methods.get(i).getcolumnName());
    }
  }

  /**
   * 读取所有的get方法并获取get方法上annotation，如果设置了excel导出，则对此read方法和annotation进行保存
   * 并按照annotation的SerialNumber的值进行排序
   * 
   * @return
   */
  private List<ExcelExportMethod> getExcelExportMethods() {
    PropertyDescriptor[] pds = PropertyUtils.getPropertyDescriptors(entityClass);

    List<ExcelExportMethod> methods = new ArrayList<ExcelExportMethod>(pds.length);
    for (PropertyDescriptor pd : pds) {
      Method readMethod = pd.getReadMethod();
      ExportField exportField = readMethod.getAnnotation(ExportField.class);
      if (null != exportField) {
        methods.add(new ExcelExportMethod(exportField, readMethod));
      }
    }
    // 对get方法进行排序（按照输出列的顺序进行排序）
    Collections.sort(methods, new Comparator<ExcelExportMethod>() {

      @Override
      public int compare(ExcelExportMethod o1, ExcelExportMethod o2) {
        return o1.getSerialNumber() - o2.getSerialNumber();
      }

    });
    return methods;
  }

  public String getExportFileName() {
    Excel excel = entityClass.getAnnotation(Excel.class);
    return excel == null ? null : excel.fileName();
  }
}
