package com.chinareny2k.commons.excel.exception;

/**
 * 当前sheet无法导入或导出Excel数据时抛出的异常
 * @author Tsunami
 *
 */
public class SkipSheetException extends RuntimeException{

	private static final long serialVersionUID = -3060770586478707387L;

	public SkipSheetException(String message){
		super(message);
	}
}
