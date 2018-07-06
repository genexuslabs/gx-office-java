package com.genexus.office.poi.xssf;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.genexus.diagnostics.core.ILogger;
import com.genexus.diagnostics.core.LogManager;
import com.genexus.gxoffice.Constants;
import com.genexus.gxoffice.IGxError;
import com.genexus.office.IExcelCellRange;
import com.genexus.office.IExcelSpreadsheet;
import com.genexus.office.IExcelWorksheet;
import com.genexus.office.exception.ExcelException;
import com.genexus.office.exception.ExcelTemplateNotFoundException;
import com.genexus.util.GXFile;

public class ExcelSpreadsheet implements IExcelSpreadsheet
{
	public static final ILogger logger = LogManager.getLogger(ExcelSpreadsheet.class);
	private XSSFWorkbook _workbook;
	private String _documentFileName;
	private boolean _autoFitColumnsOnSave = false;
	
	
	private boolean _isReadonly;
	private IGxError _errorHandler;
	
	private StylesCache _stylesCache;
	
	public ExcelSpreadsheet(IGxError errHandler, String fileName, String template) throws ExcelTemplateNotFoundException, IOException
	{
		_errorHandler = errHandler;
		if (fileName.indexOf('.') == -1)
		{
			fileName += ".xlsx";
		}

		if (!template.equals(""))
		{
			GXFile templateFile = new GXFile(template);
			if (templateFile.exists())
			{
				_workbook = new XSSFWorkbook(templateFile.getStream());
			} else
			{				
				throw new ExcelTemplateNotFoundException();
			}
		} else
		{
			GXFile file = new GXFile(fileName, Constants.EXTERNAL_PRIVATE_UPLOAD);
			if (file.exists())
			{
				_workbook = new XSSFWorkbook(file.getStream());
			} else
			{
				_workbook = new XSSFWorkbook();
			}
		}
		
		_documentFileName = fileName;

		_stylesCache = new StylesCache(_workbook);

	}

	public boolean getAutoFit() {
		return _autoFitColumnsOnSave;
	}

	public void setAutofit(boolean autoFitColumnsOnSave) {
		this._autoFitColumnsOnSave = autoFitColumnsOnSave;
	}
	
	public Boolean save() throws ExcelException
	{
		return saveAsImpl(_documentFileName);
	}

	private Boolean saveAsImpl(String fileName) throws ExcelException
	{
		ByteArrayOutputStream fs = null;
		ByteArrayInputStream in = null;
		GXFile file = null;
		
		autoFitColumns();
		recalculateFormulas();

		try
		{
			fs = new ByteArrayOutputStream();
			_workbook.write(fs);
			in = new ByteArrayInputStream(fs.toByteArray());
			fs.close();
			file = new GXFile(fileName, Constants.EXTERNAL_PRIVATE_UPLOAD);
			file.create(in, true);
			in.close();
			file.close();
		} catch (Exception e)
		{
			try
			{
				if (fs != null)
					fs.close();
				if (in != null)
					in.close();
				if (file != null)
					file.close();
			} catch (Exception e1)
			{
				logger.error("saveAsImpl", e1);
			}

			throw new ExcelException(12, "GXOffice Error: " + e.toString());
		}
		return true;
	}

	public Boolean saveAs(String newFileName) throws ExcelException
	{
		return saveAsImpl(newFileName);
	}

	public Boolean close() throws ExcelException
	{
		return save();
	}

	public IExcelCellRange getCells(IExcelWorksheet worksheet, int startRow, int startCol, int rowCount, int colCount) throws ExcelException
	{		
		return new ExcelCells(_errorHandler, this, _workbook, _workbook.getSheet(worksheet.getName()), startRow - 1, startCol - 1, rowCount, colCount, _isReadonly, _stylesCache);		
	}

	public IExcelCellRange getCell(IExcelWorksheet worksheet, int startRow, int startCol) throws ExcelException
	{
		return getCells(worksheet, startRow, startCol, 1, 1);
	}

	public Boolean insertRow(IExcelWorksheet worksheet, int rowIdx, int rowCount)
	{
		XSSFSheet sheet = getSheet(worksheet);

		int createNewRowAt = rowIdx; // Add the new row between row 9 and 10

		if (sheet != null)
		{
			for (int i = 1; i <= rowCount; i++)
			{

				int lastRow = sheet.getLastRowNum();

				XSSFRow newRow = sheet.createRow(createNewRowAt);
				sheet.shiftRows(createNewRowAt, lastRow, 1, true, false);
			}
			return true;
		}
		return false;
	}

	public Boolean insertColumn(IExcelWorksheet worksheet, int colIdx, int colCount)
	{
		/*
		 * XSSFSheet sheet = getSheet(worksheet); int createNewColumnAt = colIdx; //Add
		 * the new row between row 9 and 10
		 * 
		 * if (sheet != null) { for (int i = 1; i<= colCount; i++) {
		 * 
		 * int lastRow = sheet.getLastRowNum(); sheet.shi(createNewColumnAt, lastRow, 1,
		 * true, false); XSSFRow newRow = sheet.createRow(createNewColumnAt); } return
		 * true; } return false;
		 */
		return false; // POI not suppoerted
	}

	public Boolean deleteRow(IExcelWorksheet worksheet, int rowIdx)
	{
		XSSFSheet sheet = getSheet(worksheet);
		if (sheet != null)
		{
			XSSFRow row = sheet.getRow(rowIdx);
			if (row != null)
			{
				sheet.removeRow(row);
				return true;
			}
		}
		return false;
	}

	public List<ExcelWorksheet> getWorksheets()
	{
		List<ExcelWorksheet> list = new ArrayList<ExcelWorksheet>();
		for (int i = 0; i < _workbook.getNumberOfSheets(); i++)
		{
			XSSFSheet sheet = _workbook.getSheetAt(i);
			if (sheet != null)
			{
				list.add(new ExcelWorksheet(sheet));
			}
		}
		return list;
	}

	public Boolean insertWorksheet(String newSheetName, int idx)
	{
		XSSFSheet newSheet = null;
		if (_workbook.getSheet(newSheetName) == null)
		{
			newSheet = _workbook.createSheet(newSheetName);
			//_workbook.setSheetOrder(newSheetName, idx);
		}
		return newSheet != null;
	}

	private XSSFSheet getSheet(IExcelWorksheet sheet)
	{
		return _workbook.getSheet(sheet.getName());
	}

	private void recalculateFormulas()
	{
		try
		{
			_workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
			_workbook.setForceFormulaRecalculation(true);
		} catch (Exception e)
		{
			logger.error("recalculateFormulas", e);
		}
	}

	private void autoFitColumns()
	{
		if (_autoFitColumnsOnSave)
		{
			int sheetsCount = _workbook.getNumberOfSheets();
			for (int i = 0; i < sheetsCount; i++)
			{
				org.apache.poi.ss.usermodel.Sheet sheet = _workbook.getSheetAt(i);

				Row row = sheet.getRow(0);
				if (row != null)
				{
					int columnCount = row.getPhysicalNumberOfCells();
					for (int j = 0; j < columnCount; j++)
					{
						sheet.autoSizeColumn(j);
					}
				}
			}
		}
	}

	@Override
	public ExcelWorksheet getWorkSheet(String name) {
		XSSFSheet sheet = _workbook.getSheet(name);
		if (sheet != null)
			return new ExcelWorksheet(sheet);
		return null;
	}

	@Override
	public Boolean getAutofit() {
		return _autoFitColumnsOnSave;
	}

	@Override
	public void setColumnWidth(IExcelWorksheet worksheet, int colIdx, int width) {
		XSSFSheet sheet = _workbook.getSheet(worksheet.getName());
		if (colIdx >= 1 && sheet != null && width <= 255) {
			sheet.setColumnWidth(colIdx - 1, 256 * width);
		}
		
	}

	@Override
	public void setRowHeight(IExcelWorksheet worksheet, int rowIdx, int height) {	
		XSSFSheet sheet = _workbook.getSheet(worksheet.getName());
		if (rowIdx >=1 && sheet != null) {
			rowIdx = rowIdx - 1;
			if (sheet.getRow(rowIdx) == null)
			{
				sheet.createRow(rowIdx);
			}
			sheet.getRow(rowIdx).setHeightInPoints((short) height);
		}
	}
}
