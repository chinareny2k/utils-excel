package com.chinareny2k.commons.excel.conversion;


import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.Date;

import org.apache.commons.lang.time.DateUtils;

import com.chinareny2k.commons.excel.exception.ParseException;

public class DefaultExcelConverter implements TypeConverter {
  private static final String NULL_STRING = null;

  @Override
  public Object convertValue(Object value, Class<?> toType, String[] format) throws ParseException {
    Object result = null;
    if (null == value) {
      return null;
    } else {
      if ((toType == Integer.class) || (toType == Integer.TYPE))
        result = Integer.valueOf((int) longValue(value));
      if ((toType == Double.class) || (toType == Double.TYPE))
        result = new Double(doubleValue(value));
      if ((toType == Boolean.class) || (toType == Boolean.TYPE))
        result = booleanValue(value) ? Boolean.TRUE : Boolean.FALSE;
      if ((toType == Byte.class) || (toType == Byte.TYPE))
        result = Byte.valueOf((byte) longValue(value));
      if ((toType == Character.class) || (toType == Character.TYPE))
        result = new Character((char) longValue(value));
      if ((toType == Short.class) || (toType == Short.TYPE))
        result = Short.valueOf((short) longValue(value));
      if ((toType == Long.class) || (toType == Long.TYPE))
        result = Long.valueOf(longValue(value));
      if ((toType == Float.class) || (toType == Float.TYPE))
        result = new Float(doubleValue(value));
      if (toType == BigInteger.class)
        result = bigIntValue(value);
      if (toType == BigDecimal.class)
        result = bigDecValue(value);
      if (toType == String.class)
        result = stringValue(value);
       if(toType == Date.class)
       result = dateValue(value, format);
      return result;
    }
  }

  public static long longValue(Object value) throws ParseException {
    try {
      if (value == null)
        return 0L;
      Class<?> c = value.getClass();
      if (c.getSuperclass() == Number.class)
        return ((Number) value).longValue();
      if (c == Boolean.class)
        return ((Boolean) value).booleanValue() ? 1 : 0;
      if (c == Character.class)
        return ((Character) value).charValue();
      return Long.parseLong(stringValue(value, true));
    } catch (Exception e) {
      throw new ParseException("表格内容'" + stringValue(value) + "'无法转换成数字类型");
    }
  }

  public static double doubleValue(Object value) throws ParseException {
    try {
      if (value == null)
        return 0.0;
      Class<?> c = value.getClass();
      if (c.getSuperclass() == Number.class)
        return ((Number) value).doubleValue();
      if (c == Boolean.class)
        return ((Boolean) value).booleanValue() ? 1 : 0;
      if (c == Character.class)
        return ((Character) value).charValue();
      String s = stringValue(value, true);

      return (s.length() == 0) ? 0.0 : Double.parseDouble(s);
      /*
       * For 1.1 parseDouble() is not available
       */
      // return Double.valueOf( value.toString() ).doubleValue();
    } catch (Exception e) {
      throw new ParseException("表格内容'" + stringValue(value) + "'无法转换成小数类型");
    }
  }

  public static BigInteger bigIntValue(Object value) throws ParseException {
    try {
      if (value == null)
        return BigInteger.valueOf(0L);
      Class<?> c = value.getClass();
      if (c == BigInteger.class)
        return (BigInteger) value;
      if (c == BigDecimal.class)
        return ((BigDecimal) value).toBigInteger();
      if (c.getSuperclass() == Number.class)
        return BigInteger.valueOf(((Number) value).longValue());
      if (c == Boolean.class)
        return BigInteger.valueOf(((Boolean) value).booleanValue() ? 1 : 0);
      if (c == Character.class)
        return BigInteger.valueOf(((Character) value).charValue());
      return new BigInteger(stringValue(value, true));
    } catch (Exception e) {
      throw new ParseException("表格内容'" + stringValue(value) + "'无法转换成超大整数");
    }
  }

  public static String stringValue(Object value, boolean trim) {
    String result;

    if (value == null) {
      result = NULL_STRING;
    } else {
      result = value.toString();
      if (trim) {
        result = result.trim();
      }
    }
    return result;
  }

  public static boolean booleanValue(Object value) {
    if (value == null)
      return false;
    Class<?> c = value.getClass();
    if (c == Boolean.class)
      return ((Boolean) value).booleanValue();
    // if ( c == String.class )
    // return ((String)value).length() > 0;
    if (c == Character.class)
      return ((Character) value).charValue() != 0;
    if (value instanceof Number)
      return ((Number) value).doubleValue() != 0;
    return true; // non-null
  }

  public static BigDecimal bigDecValue(Object value) throws NumberFormatException {
    if (value == null)
      return BigDecimal.valueOf(0L);
    Class<?> c = value.getClass();
    if (c == BigDecimal.class)
      return (BigDecimal) value;
    if (c == BigInteger.class)
      return new BigDecimal((BigInteger) value);
    if (c.getSuperclass() == Number.class)
      return new BigDecimal(((Number) value).doubleValue());
    if (c == Boolean.class)
      return BigDecimal.valueOf(((Boolean) value).booleanValue() ? 1 : 0);
    if (c == Character.class)
      return BigDecimal.valueOf(((Character) value).charValue());
    return new BigDecimal(stringValue(value, true));
  }

  public static String stringValue(Object value) {
    return stringValue(value, false);
  }
  public static Date dateValue(Object value,String[] format) throws ParseException{
    try {
      if(value instanceof Date){
        return (Date)value;
      }
      return DateUtils.parseDate(stringValue(value,true), format);
    } catch (java.text.ParseException e) {
      throw new ParseException("表格内容'" + stringValue(value) + "'不符合日期格式" + Arrays.toString(format));
    }
  }
}
