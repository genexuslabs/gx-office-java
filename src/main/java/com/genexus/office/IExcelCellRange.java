package com.genexus.office;

import java.util.Date;

import com.genexus.office.style.ExcelFill;
import com.genexus.office.style.ExcelFont;
import com.genexus.office.style.ExcelStyle;

public interface IExcelCellRange
{
	public int getRowStart();

	public int getRowEnd();

	public int getColumnStart();

	public int getColumnEnd();

	public String getCellAdress();

	public String getValueType();

	/*
	 * 
	 * D: For date or datetime types C: For character type N: For numeric type U: If
	 * the type is unknown
	 */
	public String getText();

	public java.math.BigDecimal getNumericValue();

	public Date getDateValue();

	public Boolean setText(String value);

	public Boolean setNumericValue(java.math.BigDecimal value);

	public Boolean setDateValue(Date value);

	public Boolean empty();

	public Boolean mergeCells();

	public Boolean setCellStyle(ExcelStyle style);

	public ExcelStyle getCellStyle();
	
	
}
