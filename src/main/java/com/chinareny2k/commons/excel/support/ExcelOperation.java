package com.chinareny2k.commons.excel.support;

import com.chinareny2k.commons.excel.conversion.TypeConverter;

public class ExcelOperation {
	
	private TypeConverter converter;
	private boolean defaultIgnoreInconsistentSheet = true;
	
	public ExcelOperation() {
		
	}

	public TypeConverter getConverter() {
		return converter;
	}

	public void setConverter(TypeConverter converter) {
		this.converter = converter;
	}

	public boolean isDefaultIgnoreInconsistentSheet() {
		return defaultIgnoreInconsistentSheet;
	}

	public void setDefaultIgnoreInconsistentSheet(
			boolean defaultIgnoreInconsistentSheet) {
		this.defaultIgnoreInconsistentSheet = defaultIgnoreInconsistentSheet;
	}

	
	
}
