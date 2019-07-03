package com.chinareny2k.commons.excel.handler;


import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.chinareny2k.commons.excel.ExcelImportHandler;
import com.chinareny2k.commons.excel.ExcelImportMethod;
import com.chinareny2k.commons.excel.annotation.ImportField;
import com.chinareny2k.commons.excel.conversion.TypeConverter;
import com.chinareny2k.commons.excel.exception.ParseException;
import com.chinareny2k.commons.excel.exception.SkipSheetException;
import com.chinareny2k.commons.excel.support.ExcelErrorHandler;
import com.chinareny2k.commons.excel.support.ExcelErrorHandler.ErrorPosition;
import com.chinareny2k.commons.excel.support.ExcelRegionResult;
import com.chinareny2k.commons.excel.utils.ExcelUtils;

public class ExcelAnnotationImportHandler<Item> implements ExcelImportHandler<Item> {

  private final Class<Item> entityClass;
  private static Logger logger = Logger.getLogger(ExcelAnnotationImportHandler.class);
  private TypeConverter converter;
  private boolean defaultIgnoreInconsistentSheet = true;
  public ExcelErrorHandler errorHandler = new ExcelErrorHandler();

  ExcelAnnotationImportHandler(Class<Item> entityClass, TypeConverter converter) {
    this.entityClass = entityClass;
    this.converter = converter;
  }

  @Override
  public void imports(InputStream in, ExcelItemCallBack<Item> callBack) throws Exception {
	  imports(in,callBack,defaultIgnoreInconsistentSheet);
  }
  
  /**
   * 导入Excel文件
   * 
   * @param file
   *          excel 文件
   * @param validation
   *          数据校验接口
   * @return 
   * @throws APPException
   */
  @Override
	public void imports(InputStream in, ExcelItemCallBack<Item> callBack,boolean ignoreInconsistentSheet) throws Exception {

		// 获取Excel表格
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(in);
			int sheetcount = workbook.getNumberOfSheets();
			for (int k = 0; k < sheetcount; k++) {
				Sheet sheet = workbook.getSheetAt(k);
				doImports(callBack, sheet,ignoreInconsistentSheet);
			}

		} finally {
			if (workbook != null) {
				workbook.close();
			}
		}
	}

  @Override
  public List<Item> imports(InputStream in) throws Exception{
	  return imports(in,defaultIgnoreInconsistentSheet);
  }

  @Override
public List<Item> imports(InputStream in,boolean ignoreInconsistentSheet) throws Exception {
	  final List<Item> list = new ArrayList<>();
	   imports(in,new ExcelItemCallBack<Item>() {

		@Override
		public void validateData(Item data,ExcelValidationHandler handler)throws Exception {
			return;
		}

		@Override
		public void saveData(Item data) throws Exception {
			list.add(data);
		}
	},ignoreInconsistentSheet);
	   return list;
	  }
  
  private void doImports(ExcelItemCallBack<Item> callBack, Sheet sheet,boolean ignoreInconsistentSheet) throws Exception {
		// 获取Excel导入的annotation
	    Map<String, ExcelImportMethod> methodMapping = getExcelImportMethodMapping();
		try {
			// 获取Excel表头信息
			List<String> headTextList = getImportExcelHead(sheet);
			// 校验表头信息
			validateExcelHead(methodMapping, headTextList);
			// 表格内容转换
			transFormData(sheet, methodMapping, headTextList, callBack);
		} catch (SkipSheetException e) {
			if (!ignoreInconsistentSheet){
				errorHandler.addErrors(sheet.getSheetName(),0,e.getMessage());
			}
		}
	}
  
  
  /**
   * 查询需导入的method方法
   * 
   * @return
   * @throws APPException
   */
  private Map<String, ExcelImportMethod> getExcelImportMethodMapping() throws Exception {
    PropertyDescriptor[] pds = PropertyUtils.getPropertyDescriptors(entityClass);

    Map<String, ExcelImportMethod> methodMapping = new LinkedHashMap<String, ExcelImportMethod>(pds.length);
    for (PropertyDescriptor pd : pds) {
      if (!pd.getName().equals("class")) {
        Method writeMethod = pd.getWriteMethod();
        Class<?> toType = pd.getPropertyType();

        ImportField importField = writeMethod.getAnnotation(ImportField.class);
        if (null != importField) {
          methodMapping.put(importField.columnName(), new ExcelImportMethod(importField, writeMethod, toType));
        }
      }
    }
    if (MapUtils.isEmpty(methodMapping)) {
      logger.error("导入Excel失败：没有找到对应的导入项");
      throw new Exception("导入Excel失败：没有找到对应的导入项");
    }
    return methodMapping;
  }


  /**
   * 获取Excel表头信息
   * 
   * @param sheet
   * @return
   * @throws APPException
   */
  private List<String> getImportExcelHead(Sheet sheet) throws Exception {
    Row row = sheet.getRow(0);
    if (null == row || row.getPhysicalNumberOfCells() <= 0) {
      logger.error("导入Excel失败：表格为空或未找到标题信息");
      throw new SkipSheetException("导入Excel失败：表格为空或未找到标题信息");
    }

    Set<String> headSet = new LinkedHashSet<String>(row.getPhysicalNumberOfCells());

    for (int i = 0; i < row.getLastCellNum(); i++) {
      String headText = getHeadItemTitle(row, i);
      if (headSet.contains(headText)) {
        logger.error("导入Excel失败：表头第" + i + "列与其他列内容重复");
        throw new SkipSheetException("导入Excel失败：表头第" + i + "列与其他列内容重复");
      }
      headSet.add(headText);
    }

    List<String> headTextList = new ArrayList<String>(headSet.size());
    headTextList.addAll(headSet);

    return headTextList;
  }

  /**
   * 获取表头单元格的内容
   * 
   * @param row
   *          表头所在行
   * @param num
   *          单元格序号
   * @return
   * @throws APPException
   */
  private String getHeadItemTitle(Row row, int num) throws Exception {
    Cell cell = row.getCell(num);
    if (null == cell) {
      logger.error("导入Excel失败：表头第" + num + "列为空");
      throw new Exception("导入Excel失败：表头第" + num + "列为空");
    }
    String headText = cell.getStringCellValue();
    if (StringUtils.isBlank(headText)) {
      logger.error("导入Excel失败：表头第" + num + "列为空");
      throw new Exception("导入Excel失败：表头第" + num + "列为空");
    }
    return StringUtils.trim(headText);
  }

  /**
   * 校验表头，确保导入的EXCEL表格的表头和Annotation表述的一致
   * 
   * @param methodMapping
   * @param headTextList
   * @throws APPException
   */
  private void validateExcelHead(Map<String, ExcelImportMethod> methodMapping, List<String> headTextList) throws Exception {
    StringBuffer buffer = new StringBuffer();
    for (String key : methodMapping.keySet()) {
      if (!headTextList.contains(key)) {
        buffer.append("列名：" + key + "不存在；");
      }
    }
    if (StringUtils.isNotBlank(buffer.toString())) {
      String resultBuffer = buffer.toString();
      resultBuffer.substring(0, resultBuffer.length() - 1);
      logger.error("导入Excel失败：" + resultBuffer + "请按照模板进行导入");
      throw new SkipSheetException("导入Excel失败：" + resultBuffer + "请按照模板进行导入");
    }
  }

  private Item parseData(Row row, Map<String, ExcelImportMethod> methodMapping, List<String> headTextList) throws ParseException,
      Exception {
    Item data = null;
    if (null != row && 0 <= row.getPhysicalNumberOfCells()) {
      // 先遍历列名，通过列名找到列号，通过列号读取出相应的列值
      for (Entry<String, ExcelImportMethod> entry : methodMapping.entrySet()) {
        String columnName = entry.getKey();
        ExcelImportMethod method = entry.getValue();
        int columnNum = headTextList.indexOf(columnName);
        ExcelRegionResult result = ExcelUtils.isMergedRegion(row.getSheet(), row.getRowNum(), columnNum);
        Cell cell = row.getCell(columnNum);
        if(result.isMerged()){
        	cell = row.getSheet().getRow(result.getStartRow()).getCell(result.getStartCol());
        }
        // 取得Excel单元格数据
        Object cellValue = ExcelUtils.getCellValue(cell);
        if (null == cellValue) {
          continue;
        }
        try {
          data = data == null ? entityClass.newInstance() : data;
        } catch (Exception e) {
          throw new Exception("导入Excel失败：对象无法进行实例化");
        }
        // 类型转换
        try {
          Object value = converter.convertValue(cellValue, method.getToType(), method.getFormat());
          method.getWriteMethod().invoke(data, new Object[] { value });
        } catch (Exception e) {
          throw new ParseException(method.getcolumnName(), e.getMessage());
        }
      }
    }
    return data;
  }

  /**
   * 导入信息转换，将导入的信息转换成相应的可存储信息。
   * 
   * @param sheet
   *          表格
   * @param methodMapping
   *          annotation的映射
   * @param headTextList
   *          表头映射
   * @return 
   * @return
   * @throws APPException
   */
  private void transFormData(Sheet sheet, Map<String, ExcelImportMethod> methodMapping, List<String> headTextList,ExcelItemCallBack<Item> callback) throws Exception {

    boolean b = false;
    
    for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
      Row row = sheet.getRow(i);
      // 将Excel的一行数据转换成对应的类
      Item data = null;
      try {
        data = parseData(row, methodMapping, headTextList);
        if (null != data) {
          b = true;
          callback.execute(data);
          if (callback.hasError()) {
            errorHandler.addErrors(sheet.getSheetName(),i+1, callback.getErrors());
          }
        }
      } catch (ParseException e) {
        errorHandler.addErrors(sheet.getSheetName(),i+1, (StringUtils.isBlank(e.getColumnName()) ? "" : ("列:" + "'" + e.getColumnName() + "'")) + e.getMessage());
      }
    }
    if(!b && !hasErrors()) throw new SkipSheetException("Excel导入数据为空，请录入数据"); 
  }

  /**
   * 判断是否产生错误信息
   * 
   * @return
   */
  public boolean hasErrors() {
    return errorHandler.hasErrors();
  }

  /**
   * 返回所有错误信息
   * 
   * @return
   */
  public Map<ErrorPosition, List<String>> errors() {
    return errorHandler.errors();
  }


}
