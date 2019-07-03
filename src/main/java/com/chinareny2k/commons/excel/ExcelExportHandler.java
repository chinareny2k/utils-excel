package com.chinareny2k.commons.excel;


import java.io.OutputStream;
import java.util.List;


public interface ExcelExportHandler<item> {
  public static final String EXCEL_BOOLEAN_MAPPING_YES = "是";
  public static final String EXCEL_BOOLEAN_MAPPING_NO = "否";
  public static final float EXCEL_COLUMNWIDTH_PIXELCONVERSION = 36.5f;
  
  public String getExportFileName();
  public void export(List<item> list, OutputStream out) throws Exception;
}
