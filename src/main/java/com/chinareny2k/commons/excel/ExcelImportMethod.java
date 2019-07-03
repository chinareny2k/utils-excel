package com.chinareny2k.commons.excel;


import java.lang.reflect.Method;

import com.chinareny2k.commons.excel.annotation.ImportField;


public class ExcelImportMethod {
	private ImportField importField;
	private Method writeMethod;
	private Class<?> toType;

	public ExcelImportMethod(ImportField importField,Method writeMethod,Class<?> toType){
		this.importField = importField;
		this.writeMethod = writeMethod;
		this.toType = toType;
	}

	public ImportField getImportField() {
		return importField;
	}

	public void setImportField(ImportField importField) {
		this.importField = importField;
	}

	public Method getWriteMethod() {
		return writeMethod;
	}

	public void setWriteMethod(Method writeMethod) {
		this.writeMethod = writeMethod;
	}

	public String getcolumnName(){
		return this.importField.columnName();
	}


	public boolean allowBlank() {
		return this.importField.allowBlank();
	}

	public String[] getFormat() {
		return this.importField.format();
	}

	public Class<?> getToType() {
		return toType;
	}

	public void setToType(Class<?> toType) {
		this.toType = toType;
	}
}
