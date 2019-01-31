package com.genexus.office.poi.xssf;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
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
	public boolean setActiveWorkSheet(String name) {
		int idx = _workbook.getSheetIndex(name);
		if (idx >= 0)
		{
			_workbook.getSheetAt(idx).setSelected(true);
			_workbook.setActiveSheet(idx);
			_workbook.setSelectedTab(idx);
		}
		return idx >= 0;
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

	@Override
	public Boolean deleteColumn(IExcelWorksheet worksheet, int colIdx) {
		XSSFSheet sheet = _workbook.getSheet(worksheet.getName());
		if (colIdx >=0) {
			return deleteColumnImpl(sheet, colIdx);
		}
		return false;
	}
	
	private Boolean deleteColumnImpl(XSSFSheet sheet, int columnToDelete) {
	    for (int rId = 0; rId <= sheet.getLastRowNum(); rId++) {
	        Row row = sheet.getRow(rId);
	        for (int cID = columnToDelete; cID <= row.getLastCellNum(); cID++) {
	        	Cell cOld = row.getCell(cID);	        		           
	            if (cOld != null) {
	                row.removeCell(cOld);
	            }
	            Cell cNext = row.getCell(cID + 1);
	            if (cNext != null) {
	                Cell cNew = row.createCell(cID, cNext.getCellTypeEnum());
	                cloneCell(cNext, cNew);
	                //Set the column width only on the first row.
	                //Other wise the second row will overwrite the original column width set previously.
	                if(rId == 0) {
	                    sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));

	                }
	            }
	        }
	    }
	    return true;
	}

	private int getNumberOfRows(XSSFSheet sheet) {			
		int rowNum = sheet.getLastRowNum() + 1;
		return rowNum;
	}

	public int getNrColumns(XSSFSheet sheet) {		
		Row headerRow = sheet.getRow(0);
		return headerRow.getLastCellNum();		
	}
	
	public void insertNewColumnBefore(XSSFSheet sheet, int columnIndex) {
		FormulaEvaluator evaluator = _workbook.getCreationHelper()
				.createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();
		
		int nrRows = getNumberOfRows(sheet);
		int nrCols = getNrColumns(sheet);
		
		for (int row = 0; row < nrRows; row++) {
			Row r = sheet.getRow(row);

			if (r == null) {
				continue;
			}

			// shift to right
			for (int col = nrCols; col > columnIndex; col--) {
				Cell rightCell = r.getCell(col);
				if (rightCell != null) {
					r.removeCell(rightCell);
				}

				Cell leftCell = r.getCell(col - 1);

				if (leftCell != null) {
					Cell newCell = r.createCell(col, leftCell.getCellTypeEnum());
					cloneCell(newCell, leftCell);
					/*if (newCell.getCellTypeEnum() == CellType.FORMULA) {
						newCell.setCellFormula(ExcelHelper.updateFormula(newCell.getCellFormula(), columnIndex));
						evaluator.notifySetFormula(newCell);
						CellValue cellValue = evaluator.evaluate(newCell);
						evaluator.evaluateFormulaCell(newCell);						
					}*/
				}
			}

			// delete old column
			CellType cellType = CellType.BLANK;

			Cell currentEmptyWeekCell = r.getCell(columnIndex);
			if (currentEmptyWeekCell != null) {
//				cellType = currentEmptyWeekCell.getCellType();
				r.removeCell(currentEmptyWeekCell);
			}

			// create new column
			r.createCell(columnIndex, cellType);
		}

		// Adjust the column widths
		for (int col = nrCols; col > columnIndex; col--) {
			sheet.setColumnWidth(col, sheet.getColumnWidth(col - 1));
		}

		// currently updates formula on the last cell of the moved column
		// TODO: update all cells if their formulas contain references to the moved cell
//		Row specialRow = sheet.getRow(nrRows-1);
//		Cell cellFormula = specialRow.createCell(nrCols - 1);
//		cellFormula.setCellType(XSSFCell.CELL_TYPE_FORMULA);
//		cellFormula.setCellFormula(formula);

		//XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
	}
	
	/*
	 * Takes an existing Cell and merges all the styles and forumla into the new
	 * one
	 */
	private static void cloneCell(Cell cNew, Cell cOld) {
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());
	
		switch (cOld.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN: {
			cNew.setCellValue(cOld.getBooleanCellValue());
			break;
		}
		case Cell.CELL_TYPE_NUMERIC: {
			cNew.setCellValue(cOld.getNumericCellValue());
			break;
		}
		case Cell.CELL_TYPE_STRING: {
			cNew.setCellValue(cOld.getStringCellValue());
			break;
		}
		case Cell.CELL_TYPE_ERROR: {
			cNew.setCellValue(cOld.getErrorCellValue());
			break;
		}
		case Cell.CELL_TYPE_FORMULA: {
			cNew.setCellFormula(cOld.getCellFormula());
			break;
		}
		}
	}

	@Override
	public Boolean deleteSheet(int sheetIdx) {
		if (_workbook.getNumberOfSheets() > sheetIdx) {
			_workbook.removeSheetAt(sheetIdx);
			return true;
		}
		return false;			
	}

	@Override
	public Boolean deleteSheet(String sheetName) {		
		if (_workbook.getSheetIndex(sheetName) >= 0) {			
			_workbook.removeSheetAt(_workbook.getSheetIndex(sheetName));
			return true;
		}
		return false;
	}

	@Override
	public Boolean toggleColumn(IExcelWorksheet worksheet, int colIdx, Boolean visible) {
		XSSFSheet sheet = _workbook.getSheet(worksheet.getName());
		if (sheet != null) {
			sheet.setColumnHidden(colIdx, !visible);
			return true;
		}
		return false;
	}

	@Override
	public Boolean toggleRow(IExcelWorksheet worksheet, int i, Boolean visible) {
		XSSFSheet sheet = _workbook.getSheet(worksheet.getName());
		if (sheet != null) {			
			XSSFRow row = sheet.getRow(i);
			if (row != null) {
				CellStyle style = _workbook.createCellStyle();
				style.setHidden(!visible); //Does not work..
				row.setRowStyle(style);
				row.setZeroHeight(!visible);
				return true;
			}
		}
		return false;
	}
	
	
	
	/*@Override	
	public Boolean setCurrentWorksheetByName(String sheetName)
	{		
		_workbook.setActiveSheet(_workbook.getSheetIndex(sheetName));
		return true;
	}*/
	
}
