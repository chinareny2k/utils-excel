package com.chinareny2k.excel_io;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.junit.Test;

import com.chinareny2k.commons.excel.ExcelExportHandler;
import com.chinareny2k.commons.excel.ExcelImportHandler;
import com.chinareny2k.commons.excel.handler.ExcelHandlerFactory;

public class TestExcel {
	@Test
	public void testImport() throws FileNotFoundException, Exception{
		ExcelImportHandler<ResidentInfo> excelImportHandler =   ExcelHandlerFactory.getExcelImportHandler(ResidentInfo.class);
		List<ResidentInfo> list = excelImportHandler.imports( new FileInputStream(new File("d:/import.xlsx")));
		System.out.println(list);
		if (excelImportHandler.hasErrors()){
			System.out.println(excelImportHandler.errors());
		}
	}
	@Test
	public void testExport() throws FileNotFoundException, Exception{
		List<ResidentInfo> list = new ArrayList<>();
		ResidentInfo info = new ResidentInfo();
		info.setBirthday(new Date());
		info.setFcase("111");
		info.setProperty("222");
		info.setGender("333");
		info.setId("444");
		info.setMainName("555");
		info.setMedCard("666");
		info.setMoney(777f);
		info.setName("888");
		info.setfPro("999");
		list.add(info);
		ExcelExportHandler<ResidentInfo> excelExportHandler = ExcelHandlerFactory.getExcelExportHandler(ResidentInfo.class);
		excelExportHandler.export(list, new FileOutputStream(new File("d:/export.xls")));
	}
}
