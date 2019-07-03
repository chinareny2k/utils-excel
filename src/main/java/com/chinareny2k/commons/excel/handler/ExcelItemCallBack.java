package com.chinareny2k.commons.excel.handler;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.collections.CollectionUtils;
import org.apache.log4j.Logger;

public abstract class ExcelItemCallBack<Item> {
  private ExcelValidationHandler handler ;

  public ExcelItemCallBack(){
  }
  
  public abstract void validateData(Item data, ExcelValidationHandler handler) throws Exception;

  public abstract void saveData(Item data) throws Exception;

  private static Logger logger = Logger.getLogger(ExcelItemCallBack.class);

  void execute(Item data) {
    try {
      this.handler = new ExcelValidationHandler(); 
      validateData(data, handler);
      if (hasError()) {
        return;
      } else {
        saveData(data);
      }
    } catch (Exception e) {
      handler.addError(e.getMessage());
      logger.error(e);
    }
  }

  boolean hasError() {
    return CollectionUtils.isNotEmpty(handler.errorMessage);
  }
  List<String> getErrors(){
    return handler.errorMessage;
  }

   public class ExcelValidationHandler {
     private List<String> errorMessage = new ArrayList<String>();
     private ExcelValidationHandler(){};

    public void addError(String field, String message) {
      errorMessage.add(message);
    }

    public void addError(String message) {
      errorMessage.add(message);
    }

  }
}
