package com.chinareny2k.commons.excel;


import java.io.InputStream;
import java.util.List;
import java.util.Map;

import com.chinareny2k.commons.excel.handler.ExcelItemCallBack;
import com.chinareny2k.commons.excel.support.ExcelErrorHandler.ErrorPosition;



public interface ExcelImportHandler<Item> {
  public void imports(InputStream in, ExcelItemCallBack<Item> callBack,boolean ignoreInconsistentSheet) throws Exception;
  public void imports(InputStream in, ExcelItemCallBack<Item> callBack) throws Exception;
  public List<Item> imports(InputStream in) throws Exception;
  public List<Item> imports(InputStream in, boolean ignoreInconsistentSheet) throws Exception;
//  public void export(List<item> list, OutputStream out) throws Exception;
  /**
   * 判断是否产生错误信息
   * @return
   */
  public boolean hasErrors();

  /**
   * 返回所有错误信息
   * @return
   */
  public Map<ErrorPosition, List<String>> errors();

}
