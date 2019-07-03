package com.chinareny2k.commons.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
/**
 * 用来处理Excel导入字段的注解信息
 * 
 * @author Liying
 * @version $Revision$ $Date$
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
public @interface ImportField {
  public String columnName(); // 该值对应的列名

  public String[] format() default "{yyyy-MM-dd HH:mm:ss}";// 日期型字符串的格式化模板
  /**
   * 是否允许为空,默认为true，允许为空
   * 
   * @return true :允许为空 false :不允许为空 默认为true
   */
  public boolean allowBlank() default true;
  
}
