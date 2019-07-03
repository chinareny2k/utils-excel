package com.chinareny2k.commons.excel.exception;


public class ParseException extends RuntimeException {

  private static final long serialVersionUID = -6929211379747954686L;
  private int rowIndex;
  private String columnName;

  public ParseException(String message) {
    super(message);
  }

  public ParseException(int rowIndex, String columnName, String message) {
    super(message);
    this.rowIndex = rowIndex;
    this.columnName = columnName;
  }
  public ParseException(String columnName,String message){
    super(message);
    this.columnName = columnName;
  }

  public int getRowIndex() {
    return rowIndex;
  }

  public String getColumnName() {
    return columnName;
  }

}
