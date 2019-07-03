package com.chinareny2k.commons.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用来处理Excel导出字段的注解信息
 * 
 * @author Song Liang
 * @version $Revision$ $Date$
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
public @interface ExportField {
  public String columnName(); // 该值对应的列名

  public int columnWidth() default 100; // 该列的宽度

  public String format() default "yyyy-MM-dd HH:mm:ss";// 日期型字符串的格式化模板
  
  public int serialNumber();//序号
  
  /**
   * 用来处理值对应的映射
   * 如从数据库中取出的值为1，但在导出到excel的时候实际的导出信息为“否”时则需要此映射
   * 使用示例
   * mapping="1$$是||0$$否||default$$是",中间不能使用空格,其中default为默认值
   * 使用||分割每个值$$左边表示的是值，右边表示的是映射结果
   * @return
   */
  public String mapping()default "";

}
