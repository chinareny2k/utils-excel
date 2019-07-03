package com.chinareny2k.commons.excel;


import java.lang.reflect.Method;

import com.chinareny2k.commons.excel.annotation.ExportField;


public class ExcelExportMethod {
  private ExportField exportField;
  private Method readMethod;
  
  public ExcelExportMethod(ExportField exportField,Method readMethod){
    this.exportField = exportField;
    this.readMethod = readMethod;
  }

  public ExportField getExportField() {
    return exportField;
  }

  public void setExportField(ExportField exportField) {
    this.exportField = exportField;
  }

  public Method getReadMethod() {
    return readMethod;
  }

  public void setReadMethod(Method readMethod) {
    this.readMethod = readMethod;
  }
  
  public String getcolumnName(){
    return this.exportField.columnName();
  }
  
  public int getColumnWidth(){
    return this.exportField.columnWidth();
  }
  
  public String getFormat(){
    return this.exportField.format();
  }
  
  public int getSerialNumber(){
    return this.exportField.serialNumber();
  }
  
  public String getMapping(){
    return this.exportField.mapping();
  }
}
