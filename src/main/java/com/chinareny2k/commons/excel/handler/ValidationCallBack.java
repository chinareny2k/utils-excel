package com.chinareny2k.commons.excel.handler;


import com.chinareny2k.commons.excel.exception.ParseException;

public interface ValidationCallBack<Item> {
 void validateData(Item data)throws ParseException;

}
