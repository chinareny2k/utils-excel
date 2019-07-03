package com.chinareny2k.commons.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用来标识excel文件名和sheet名称的注解
 * 
 * @author Song Liang
 * @version $Revision$ $Date$
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Excel {
  public String fileName();//导出Excel文件的文件名称
  public String sheetName();//导出Excel文件的默认页名称
}
