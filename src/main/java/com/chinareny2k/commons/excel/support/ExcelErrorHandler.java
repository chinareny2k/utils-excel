package com.chinareny2k.commons.excel.support;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.MapUtils;
/**
 * 导入Excel错误类
 * @author Li Ying
 *
 */
public class ExcelErrorHandler {
	private Map<ErrorPosition, List<String>> errors = new HashMap<ErrorPosition,List<String>>();
	
	/**
	 * 添加错误信息
	 * @param sheetName excel Sheet名称
	 * @param rowIndex  当前行号
	 * @param message 错误信息
	 */
	public void addErrors(String sheetName, Integer rowIndex, String message) {
		ErrorPosition position = new ErrorPosition(sheetName, rowIndex);
	    if(errors.containsKey(position)){
	      errors.get(position).add(message);
	    }else{
	      List<String> list = new ArrayList<String>();
	      list.add(message);
	      errors.put(position, list);
	    }
	  }
	
	public void addErrors(String sheetName, Integer rowIndex,List<String> message){
	  ErrorPosition position = new ErrorPosition(sheetName, rowIndex);
	  if(errors.containsKey(position)){
	   errors.get(position).addAll(message);
	  }else{
       errors.put(position, message);
	  }
	}
	
	public boolean hasErrors(){
	  return MapUtils.isNotEmpty(errors);
	}
	 
	public Map<ErrorPosition, List<String>> errors(){
	  return Collections.unmodifiableMap(errors);
	}
	
	public class ErrorPosition{
		private String sheetName;
		private Integer rowIndex;
		private ErrorPosition(String sheetName,Integer rowIndex){
			this.sheetName = sheetName;
			this.rowIndex = rowIndex;
		}
		public String getSheetName() {
			return sheetName;
		}
		public void setSheetName(String sheetName) {
			this.sheetName = sheetName;
		}
		public Integer getRowIndex() {
			return rowIndex;
		}
		public void setRowIndex(Integer rowIndex) {
			this.rowIndex = rowIndex;
		}
		@Override
		public int hashCode() {
			final int prime = 31;
			int result = 1;
			result = prime * result + getOuterType().hashCode();
			result = prime * result + rowIndex;
			result = prime * result
					+ ((sheetName == null) ? 0 : sheetName.hashCode());
			return result;
		}
		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (obj == null)
				return false;
			if (getClass() != obj.getClass())
				return false;
			ErrorPosition other = (ErrorPosition) obj;
			if (!getOuterType().equals(other.getOuterType()))
				return false;
			if (rowIndex != other.rowIndex)
				return false;
			if (sheetName == null) {
				if (other.sheetName != null)
					return false;
			} else if (!sheetName.equals(other.sheetName))
				return false;
			return true;
		}
		private ExcelErrorHandler getOuterType() {
			return ExcelErrorHandler.this;
		}
		
	}
	
}
