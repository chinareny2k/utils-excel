package com.chinareny2k.commons.excel.conversion;

import com.chinareny2k.commons.excel.exception.ParseException;



public interface TypeConverter {
/**
 * 
 * @param value
 * @param toType
 * @param format
 * @return
 * @throws ParseException
 * @throws com.telezone.ibs.singularity.commons.excel.exception.ParseException 
 */
public Object convertValue(Object value, Class<?> toType,String[] format) throws ParseException;
}
