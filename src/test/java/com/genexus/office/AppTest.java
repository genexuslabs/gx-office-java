package com.genexus.office;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.*;

import com.genexus.office.poi.xssf.ExcelCells;
import com.genexus.office.poi.xssf.ExcelWorksheet;
import com.genexus.office.style.ExcelAlignment;
import com.genexus.office.style.ExcelStyle;

/**
 * Unit test for simple App.
 */

public class AppTest
{
	private static String basePath = "C:\\temp\\excel\\";

	/**
	 * Rigourous Test :-)
	 */
	
	@Test
	public void testjapan1()
	{
		String excel1 = basePath + "test_japan1";
		deletefile(excel1 + ".xlsx");
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel1);
		excel.setColumnWidth(1,  100);
		excel.getCells(2, 1, 1, 5).setNumericValue(new java.math.BigDecimal(123.456));
		ExcelStyle newCellStyle = new ExcelStyle();
		newCellStyle.getCellFont().setBold(true);
		excel.getCells(2, 1, 1, 5).setCellStyle(newCellStyle);
	
		excel.save();

	}
	
	@Test
	public void testjapan2()
	{
		String excel1 = basePath + "test_japan2";
		deletefile(excel1 + ".xlsx");
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel1);
		excel.setColumnWidth(1,  100);
		excel.getCells(2, 1, 5, 5).setNumericValue(new java.math.BigDecimal(123.456));
		ExcelStyle newCellStyle = new ExcelStyle();
		newCellStyle.getCellFont().setBold(true);
		excel.getCells(2, 1, 3, 3).setCellStyle(newCellStyle);
	
		excel.save();

	}
	
	@Test
	public void testActiveWorksheet()
	{
		String excel1 = basePath + "ActiveWorksheet";
		deletefile(excel1 + ".xlsx");
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel1);
		
		excel.getCells(2, 1, 5, 5).setNumericValue(new java.math.BigDecimal(123.456));
		excel.insertSheet("test1");
	
		excel.insertSheet("test2");
		excel.insertSheet("test3");
		excel.setCurrentWorksheetByName("test2");
		excel.getCells(2, 1, 5, 5).setNumericValue(new java.math.BigDecimal(3));
		excel.save();

	}
	
	
	@Test
	public void testWithoutExtensions()
	{
		String excel1 = basePath + "test_withoutextension";
		deletefile(excel1 + ".xlsx");
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel1);
		excel.insertSheet("genexus0");
		excel.insertSheet("genexus1");
		excel.insertSheet("genexus2");

		List<ExcelWorksheet> wSheets = excel.getWorksheets();
		assertTrue(wSheets.size() == 4);
		assertTrue(wSheets.get(1).getName() == "genexus0");
		assertTrue(wSheets.get(2).getName() == "genexus1");
		assertTrue(wSheets.get(3).getName() == "genexus2");

		excel.save();

	}
	
	@Test
	public void testXLSExtension()
	{
		String excel1 = basePath + "testXLSExtension.xls";
		deletefile(excel1);
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		
		excel.open(excel1);
		assertTrue(excel.getErrCode() > 0);
		excel.insertSheet("genexus0");
		excel.insertSheet("genexus1");
		excel.insertSheet("genexus2");

		List<ExcelWorksheet> wSheets = excel.getWorksheets();
		
		excel.save();
		excel.close();
	}
	
	@Test
	public void testInsertSheet()
	{
		String excel1 = basePath + "test_insert_sheet.xlsx";
		deletefile(excel1);
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel1);
		excel.insertSheet("genexus0");
		excel.insertSheet("genexus1");
		excel.insertSheet("genexus2");

		List<ExcelWorksheet> wSheets = excel.getWorksheets();
		assertTrue(wSheets.size() == 4);
		assertTrue(wSheets.get(1).getName() == "genexus0");
		assertTrue(wSheets.get(2).getName() == "genexus1");
		assertTrue(wSheets.get(3).getName() == "genexus2");

		excel.save();

	}

	@Test
	public void testSetCellValues()
	{
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		String excel2 = basePath + "excel2.xlsx";
		excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel2);
		excel.setAutofit(true);
		excel.getCells(1, 1, 1, 1).setNumericValue(new java.math.BigDecimal(100));		
		excel.getCells(2, 1, 1 , 1).setText("hola!");
		excel.getCells(3, 1, 1, 1).setDateValue(new Date());
		excel.getCells(4, 1, 1, 1).setNumericValue(new java.math.BigDecimal(66.78));

		excel.save();
		excel.close();
		// Verify previous Excel Document
		excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel2);

		assertEquals(100, excel.getCells(1, 1, 1, 1).getNumericValue().intValue());

		assertEquals("No Coindicen", excel.getCells(2, 1, 1, 1).getText(), "hola!");
		excel.save();
	}
	
	@Test
	public void testFormulas()
	{
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		String excel2 = basePath + "excel2.xlsx";
		excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel2);
		excel.setAutofit(true);
		excel.getCell(1, 1).setNumericValue(new java.math.BigDecimal(5));
		excel.getCell(2, 1).setNumericValue(new java.math.BigDecimal(6));
		excel.getCell(3, 1).setText("=A1+A2");
		excel.save();
		excel.close();
		// Verify previous Excel Document
		excel = new ExcelSpreadsheetGXWrapper();
		excel.open(excel2);

		assertEquals(11, excel.getCell(3, 1).getNumericValue().intValue());
	
		excel.save();
	}
	
	
	@Test
	public void testReadExcelFile() {
		testSetCellValues();
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		 //Test opening Existing Excel Sheet
        String excel3 = basePath + "readExcelFileTest1.xlsx";  
        excel = new ExcelSpreadsheetGXWrapper();
        excel.open(excel3);   
              
        assertEquals(excel.getCell(2, 2).getNumericValue().intValue(), 100, 0 );
        assertEquals(excel.getCell(3, 3).getNumericValue().intValue(), 100, 0);
        excel.close();
	}
	
	@Test
	public void testExcelCellRange() {
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();
		 //Test opening Existing Excel Sheet
        String excel3 = basePath + "readExcelFileTest1.xlsx";  
        excel = new ExcelSpreadsheetGXWrapper();
        excel.open(excel3);   
        
        IExcelCellRange cellRange = excel.getCells(2, 2, 5, 10);
        
        assertEquals(2, cellRange.getColumnStart(), 0 );
        assertEquals(11, cellRange.getColumnEnd(), 0 );
        assertEquals(2, cellRange.getRowStart(), 0 );
        assertEquals(6, cellRange.getRowEnd(), 0 );        
        excel.close();
	}

	
	@Test
	@Ignore
	public void testCellRangeCellAddres() {
		//Pending Implementation..
	}
	
	
	@Test
	public void testSetCurrentWorksheetByName() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_test_setCurrentWorksheetByName.xlsx";			              
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       excel.insertSheet("hoja1");
	       excel.insertSheet("hoja2");
	       excel.insertSheet("hoja3");
	       excel.save();
	       excel.close();
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       excel.setCurrentWorksheetByName("hoja2");
	       assertEquals("hoja2", excel.getCurrentWorksheet().getName() );
	       excel.getCell(5, 5).setText("hola");
	       excel.save();
	       excel.close();
	       
	       
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       excel.setCurrentWorksheetByName("hoja2");
	       assertEquals("hola", excel.getCell(5, 5).getText());
	       
	       excel.setCurrentWorksheetByName("hoja1");	       
	       assertEquals("", excel.getCell(5, 5).getText());
	      // excel.close();
	}
	
	
	@Test
	public void testCopySheet() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_testCopySheet.xlsx";	
		   deletefile(excelPath);
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       excel.insertSheet("hoja1");
	       excel.setCurrentWorksheetByName("hoja1");
	       excel.getCells(1, 1, 3, 3).setText("test");
	       excel.insertSheet("hoja2");
	       excel.insertSheet("hoja3");
	       excel.save();
	       excel.close();
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	      excel.setCurrentWorksheetByName("hoja1");
	      excel.getCurrentWorksheet().copy("hoja1Copia");
	      	excel.save();
	       excel.close();
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       excel.setCurrentWorksheetByName("hoja1Copia");
	       assertEquals("No Coindicen",excel.getCells(1, 1, 3, 3).getText(), "test");
	       excel.close();
	}
	
	@Test
	public void testgetWorksheets() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_test_getWorksheets.xlsx";			              
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       excel.insertSheet("hoja1");
	       excel.insertSheet("hoja2");
	       excel.insertSheet("hoja3");
	       excel.insertSheet("hoja4");
	       excel.save();
	       excel.close();
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       List<ExcelWorksheet> sheets = excel.getWorksheets();
	       assertEquals("hoja1", sheets.get(1).getName() );
	       assertEquals("hoja2", sheets.get(2).getName() );
	       assertEquals("hoja3", sheets.get(3).getName() );
	       assertEquals("hoja4", sheets.get(4).getName() );	       	       
	       excel.close();	       	       	   
	}
	
	@Test
	public void testProtectSheet() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_test_protectedsheet.xlsx";			              
	       excel = new ExcelSpreadsheetGXWrapper();   
	       deletefile(excelPath);
	         
	       excel.open(excelPath);
	       excel.setAutofit(true);
	       excel.insertSheet("hoja1");
	       excel.setCurrentWorksheetByName("hoja1");
	       excel.getCurrentWorksheet().setProtected("password");
	       excel.getCells(1, 1, 3, 3).setText("texto no se puede editar");
	       ExcelStyle style = new ExcelStyle();
	       style.setLocked(true);
	       excel.getCells(1, 1, 3, 3).setCellStyle(style);
	       
	       
	       ExcelCells cells = excel.getCells(5, 1, 3, 3);
	       cells.setText("texto SI se puede editar");	  
	       style = new ExcelStyle();
	       style.setLocked(false);
	       cells.setCellStyle(style);
	       excel.save();
	       excel.close();
	}
	
	@Test
	public void testgetWorksheetRename() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_test_worksheetRename.xlsx";			
		   deletefile(excelPath);
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       excel.getCurrentWorksheet().rename("defaultsheetrenamed");
	       excel.insertSheet("hoja1");
	       excel.insertSheet("hoja2");
	       excel.insertSheet("hoja3");
	       excel.insertSheet("hoja4");
	       
	       excel.save();
	       excel.close();
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       excel.getWorksheets().get(3).rename("modificada");
	       excel.save();
	       excel.close();
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       List<ExcelWorksheet> sheets = excel.getWorksheets();
	       assertEquals("hoja1", sheets.get(1).getName() );
	       assertEquals("hoja2", sheets.get(2).getName() );
	       assertEquals("modificada", sheets.get(3).getName() );
	       assertEquals("hoja4", sheets.get(4).getName() );	       	       
	       excel.close();	       	       	   
	}
	
	@Test
	public void testMergeCells() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_testMergeCells.xlsx";		
		   deletefile(excelPath);
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       excel.getCells(2, 2, 3, 3).mergeCells();
	       excel.getCells(2, 2, 3, 3).setText("merged cells");
	       excel.save();
	       excel.close();
	              	       	  
	}
	
	@Test
	public void testColumnAndRowHeight() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_testColumnAndRowHeight.xlsx";		
		   deletefile(excelPath);
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       //excel.setAutofit(false);
	       excel.getCells(1, 1, 5, 5).setText("texto de las celdas largo");
	       excel.setRowHeight(2, 50);
	       excel.setColumnWidth(1, 100);
	       excel.save();
	       excel.close();
	              	       	  
	}
	
	@Test
	public void testAlignment() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_testAlignment.xlsx";
		   deletefile(excelPath);
		   
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       //excel.getCells(2, 2, 3, 3).mergeCells();
	       excel.getCells(2, 2, 3, 3).setText("a");
	       ExcelStyle style = new ExcelStyle();
	       style.getCellAlignment().setHorizontalAlignment(ExcelAlignment.HORIZONTAL_ALIGN_RIGHT); //center
	       style.getCellAlignment().setVerticalAlignment(ExcelAlignment.VERTICAL_ALIGN_MIDDLE); //middle
	       excel.getCells(2, 2, 3, 3).setCellStyle(style);
	       excel.save();
	       excel.close();
	              	       	  
	}
	
	/*
	@Test
	public void testSetCurrentWorksheetByIndex() {
		  ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		   String excelPath = basePath + "excel_test_setCurrentWorksheetByName.xlsx";			              
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       excel.insertSheet("hoja1");
	       excel.insertSheet("hoja2");
	       excel.insertSheet("hoja3");
	       excel.save();
	       excel.close();
	       excel = new ExcelSpreadsheetGXWrapper();   
	       excel.open(excelPath);
	       
	       excel.getCell(5, 5).setText("hola");
	       excel.save();
	       excel.close();
	}
	*/
	
	@Test	
	public void testExcelCellStyle() {
	   ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
	   String excelPath = basePath + "excelStyleTest.xlsx";  
	   deletefile(excelPath);      
       excel = new ExcelSpreadsheetGXWrapper();   
       excel.open(excelPath);
       IExcelCellRange cells = excel.getCells(1, 1, 2, 2);
       
       ExcelStyle style = new ExcelStyle();
       
       cells.setText("texto1");
       style.getCellAlignment().setHorizontalAlignment(3);
       style.getCellFont().setBold(true);
       style.getCellFont().setStrike(true);
       style.getCellFont().setItalic(true);
       style.getCellFont().setSize(30);
       style.getCellFont().getColor().setColorRGB(1,1,1);
       //style.getCellFont().getColor().setColorARGB(0, 50, 100, 180);
       style.getCellFill().getCellBackColor().setColorRGB(210,180,140);       
       cells.setCellStyle(style);
    
       
              
       cells = excel.getCells(5, 2, 4, 4);
       
       cells.setText("texto2");
       style = new ExcelStyle();             
       style.getCellFont().setSize(10);
       style.getCellFont().getColor().setColorRGB(255,255,255);
       //style.getCellFont().getColor().setColorARGB(0, 50, 100, 180);
       style.getCellFill().getCellBackColor().setColorRGB(90,90,90);       
  
       cells.setCellStyle(style);
       
       
       cells = excel.getCells(10, 2, 2,2);       
       cells.setText("texto3");
       style = new ExcelStyle();
       style.getCellFont().setBold(false);
       style.getCellFont().setSize(10);
       style.getCellFont().getColor().setColorRGB(180,180,180);
       //style.getCellFont().getColor().setColorARGB(0, 50, 100, 180);
       style.getCellFill().getCellBackColor().setColorRGB(45,45,45);         
       cells.setCellStyle(style);
       
       
       excel.save();
       excel.close();
       
	}
	
	@Test
	public void testSaveAs() {
		ExcelSpreadsheetGXWrapper excel = new ExcelSpreadsheetGXWrapper();		
		 String excel3 = basePath + "readExcelFileTest1.xlsx";  
		
                 
        excel = new ExcelSpreadsheetGXWrapper();   
        excel.open(excel3);
        excel.getCells(1, 1, 15, 15).setNumericValue(new BigDecimal(100));
        String excelNew = basePath + "readExcelFileTest1new.xlsx";
        excel.saveAs(excelNew);
        excel.close();
        assertEquals(new File(excelNew).exists(), true);
        
	}

	
	private void deletefile(String path)
	{
		
		File file = new File(path);
		if (file.exists())
			file.delete();
		
	}
}
