package com.chinareny2k.excel_io;

import java.util.Date;

import com.chinareny2k.commons.excel.annotation.ExportField;
import com.chinareny2k.commons.excel.annotation.ImportField;

public class ResidentInfo {
	private String mainName;
	private String medCard;
	private Float money;
	private String fPro;
	private String name;
	private String gender;
	private String fcase;
	private String id;
	private Date birthday;
	private String property;
	@ExportField(columnName="户主姓名",serialNumber=1)
	public String getMainName() {
		return mainName;
	}

	@ImportField(allowBlank = false, columnName = "户主姓名")
	public void setMainName(String mainName) {
		this.mainName = mainName;
	}
	@ExportField(columnName="医疗证号",serialNumber=2)
	public String getMedCard() {
		return medCard;
	}

	@ImportField(allowBlank = false, columnName = "医疗证号")
	public void setMedCard(String medCard) {
		this.medCard = medCard;
	}
	@ExportField(columnName="家庭账户余额",serialNumber=3)
	public Float getMoney() {
		return money;
	}

	@ImportField(allowBlank = false, columnName = "家庭账户余额")
	public void setMoney(Float money) {
		this.money = money;
	}
	@ExportField(columnName="户属性",serialNumber=4)
	public String getfPro() {
		return fPro;
	}

	@ImportField(allowBlank = false, columnName = "户属性")
	public void setfPro(String fPro) {
		this.fPro = fPro;
	}
	@ExportField(columnName="中文姓名",serialNumber=5)
	public String getName() {
		return name;
	}

	@ImportField(allowBlank = false, columnName = "中文姓名")
	public void setName(String name) {
		this.name = name;
	}
	@ExportField(columnName="性别",serialNumber=6)
	public String getGender() {
		return gender;
	}

	@ImportField(allowBlank = false, columnName = "性别")
	public void setGender(String gender) {
		this.gender = gender;
	}
	@ExportField(columnName="家庭关系",serialNumber=7)
	public String getFcase() {
		return fcase;
	}

	@ImportField(allowBlank = false, columnName = "家庭关系")
	public void setFcase(String fcase) {
		this.fcase = fcase;
	}
	@ExportField(columnName="身份证号码",serialNumber=8)
	public String getId() {
		return id;
	}

	@ImportField(allowBlank = false, columnName = "身份证号码")
	public void setId(String id) {
		this.id = id;
	}
	@ExportField(columnName="出生日期",serialNumber=9,format="yyyy-MM-dd")
	public Date getBirthday() {
		return birthday;
	}
	@ImportField(allowBlank = false, columnName = "出生日期",format={"yyyy-MM-dd"})
	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}
	@ExportField(columnName="参合属性",serialNumber=9)
	public String getProperty() {
		return property;
	}
	@ImportField(allowBlank=false, columnName="参合属性")
	public void setProperty(String property) {
		this.property = property;
	}

	@Override
	public String toString() {
		return "ResidentInfo [mainName=" + mainName + ", medCard=" + medCard
				+ ", money=" + money + ", fPro=" + fPro + ", name=" + name
				+ ", gender=" + gender + ", fcase=" + fcase + ", id=" + id
				+ ", birthday=" + birthday + ", property=" + property + "]";
	}

}
