package com.genexus.office.poi.xssf;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.genexus.office.IExcelWorksheet;

public class ExcelWorksheet implements IExcelWorksheet
{
	private XSSFSheet _sheet;

	public ExcelWorksheet()
	{
	
	}
	
	
	public ExcelWorksheet(XSSFSheet sheet)
	{
		_sheet = sheet;
	}

	public String getName()
	{
		return _sheet.getSheetName();
	}

	public Boolean isHidden()
	{
		return false;
	}

	public Boolean rename(String newName)
	{
		if (_sheet != null) {
			XSSFWorkbook wb = _sheet.getWorkbook();
			wb.setSheetName(wb.getSheetIndex(getName()), newName);
			return getName().equals(newName);
		}
		return false;
	}


	@Override
	public Boolean copy(String newName) {
		if (_sheet != null) {
			XSSFWorkbook wb = _sheet.getWorkbook();
			wb.cloneSheet(wb.getSheetIndex(getName()), newName);
			return true;
		}
		return false;
	}

	@Override
	public void setProtected(String password) {
		if (_sheet != null) {
			if (password.length() == 0)
				_sheet.protectSheet(null);
			else
				_sheet.protectSheet(password);		
		}		
	}

}
