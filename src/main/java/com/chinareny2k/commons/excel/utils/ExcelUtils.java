package com.chinareny2k.commons.excel.utils;


import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import com.chinareny2k.commons.excel.support.ExcelRegionResult;

public class ExcelUtils {
 // private static final String template = "com/telezone/app/singularity/commons/excel/vm/ErrorMessage.vm";
  private ExcelUtils() {
  }

  
  public static Object getCellValue(Cell cell) {
    Object fieldValue = null;
    if (cell != null) {
      switch (cell.getCellTypeEnum()) {
      case STRING:
        fieldValue = StringUtils.isBlank(cell.getStringCellValue()) ? null : cell.getStringCellValue();
        break;
      case BOOLEAN:
        fieldValue = cell.getBooleanCellValue();
        break;
      case NUMERIC:
    	if (HSSFDateUtil.isCellDateFormatted(cell)) {    
    	    fieldValue = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());    
    	}else{
    		fieldValue = cell.getNumericCellValue();
    	}
        break;
      case BLANK:
        fieldValue = null;
        break;
      case ERROR:
        fieldValue = cell.getErrorCellValue();
        break;
      case FORMULA:
        fieldValue = cell.getCellFormula();
        break;
	default:
		fieldValue = null;
		break;
      }
    }
    return fieldValue;
  }
/*  public static String showErrorMessage(Map<Integer, List<String>> errorMap){
    VelocityEngine ve = new VelocityEngine();
    ve.setProperty(Velocity.RESOURCE_LOADER, "class");
    ve.setProperty("class.resource.loader.class","org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader");
    Template t = null;
    try {
      ve.init();
      t = ve.getTemplate(template,"UTF-8");
    } catch (Exception e) {
      e.printStackTrace();
    }
    
    VelocityContext context = new VelocityContext();
    context.put("errorMap", errorMap);
    //设置输出
    StringWriter writer = new StringWriter();
    //将环境数据化输出
    try {
      t.merge(context, writer);
    } catch (Exception e) {
      e.printStackTrace();
    }
    return writer.toString();
  }*/
  
  public static  ExcelRegionResult isMergedRegion(Sheet sheet,int row ,int column) {   
      int sheetMergeCount = sheet.getNumMergedRegions();   
      for (int i = 0; i < sheetMergeCount; i++) {   
            CellRangeAddress range = sheet.getMergedRegion(i);   
            int firstColumn = range.getFirstColumn(); 
            int lastColumn = range.getLastColumn();   
            int firstRow = range.getFirstRow();   
            int lastRow = range.getLastRow();   
            if(row >= firstRow && row <= lastRow){ 
                if(column >= firstColumn && column <= lastColumn){ 
               return new ExcelRegionResult(true,firstRow,lastRow,firstColumn,lastColumn);   
                } 
            }
      } 
      return new ExcelRegionResult(false,0,0,0,0);  
    } 
}
