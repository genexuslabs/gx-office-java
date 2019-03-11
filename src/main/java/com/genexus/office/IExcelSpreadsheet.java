package com.genexus.office;

import java.util.List;

import com.genexus.office.exception.ExcelException;
import com.genexus.office.poi.xssf.ExcelWorksheet;

public interface IExcelSpreadsheet
{
	// General Methods
	public Boolean save() throws ExcelException;

	public Boolean saveAs(String newFileName) throws ExcelException;

	public Boolean close() throws ExcelException;

	// CellMethods
	public IExcelCellRange getCells(IExcelWorksheet worksheet, int startRow, int startCol, int rowCount, int colCount) throws ExcelException;

	public IExcelCellRange getCell(IExcelWorksheet worksheet, int startRow, int startCol) throws ExcelException;

	public Boolean insertRow(IExcelWorksheet worksheet, int rowIdx, int rowCount);

	public Boolean deleteRow(IExcelWorksheet worksheet, int rowIdx);

	// Columns not supported
	// public Boolean insertColumn(IExcelWorksheet worksheet, int rowIdx, int
	// colIdx);
	public Boolean deleteColumn(IExcelWorksheet worksheet, int colIdx);

	// Worksheets
	public List<ExcelWorksheet> getWorksheets();
	public ExcelWorksheet getWorkSheet(String name);

	public Boolean insertWorksheet(String newSheetName, int idx);
	public Boolean getAutofit();
	public void setAutofit(boolean autofit);

	public void setColumnWidth(IExcelWorksheet worksheet, int colIdx, int width);
	public void setRowHeight(IExcelWorksheet worksheet, int rowIdx, int height);

	boolean setActiveWorkSheet(String name);

	public Boolean deleteSheet(int sheetIdx);

	public Boolean deleteSheet(String sheetName);

	public Boolean toggleColumn(IExcelWorksheet worksheet, int colIdx, Boolean visible);

	public Boolean toggleRow(IExcelWorksheet _currentWorksheet, int i, Boolean visible);


	
}
