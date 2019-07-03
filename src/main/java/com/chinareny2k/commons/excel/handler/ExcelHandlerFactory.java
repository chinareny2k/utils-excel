package com.chinareny2k.commons.excel.handler;


import java.util.HashMap;
import java.util.Map;

import com.chinareny2k.commons.excel.ExcelExportHandler;
import com.chinareny2k.commons.excel.ExcelImportHandler;
import com.chinareny2k.commons.excel.conversion.DefaultExcelConverter;
import com.chinareny2k.commons.excel.conversion.TypeConverter;

public class ExcelHandlerFactory {

  @SuppressWarnings("rawtypes")
private static  Map<Class, ExcelExportHandler> EXPORTHANDLER = new HashMap<Class, ExcelExportHandler>();

  private ExcelHandlerFactory() {
  }

  @SuppressWarnings("unchecked")
  public static <T> ExcelExportHandler<T> getExcelExportHandler(Class<T> clazz) {
    if (EXPORTHANDLER.containsKey(clazz)) {
      return EXPORTHANDLER.get(clazz);
    } else {
      ExcelAnnotationExportHandler<T> handler = new ExcelAnnotationExportHandler<T>(clazz);
      EXPORTHANDLER.put(clazz, handler);
      return handler;
    }
  }
  
  public static<Item> ExcelImportHandler<Item> getExcelImportHandler(Class<Item> entityClass){
    return new ExcelAnnotationImportHandler<Item>(entityClass,new DefaultExcelConverter());
  }
  
  public static <Item> ExcelImportHandler<Item> getExcelImportHandler(Class<Item> entityClass,TypeConverter converter){
    return new ExcelAnnotationImportHandler<Item>(entityClass,converter);
  }
}
