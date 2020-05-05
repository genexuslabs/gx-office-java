package com.genexus.office;

import com.genexus.office.exception.ExcelDocumentNotSupported;
import com.genexus.office.exception.ExcelTemplateNotFoundException;

import java.io.IOException;

public class ExcelFactory
{

	public static IExcelSpreadsheet create(IGXError handler, String filePath, String template)
			throws ExcelTemplateNotFoundException, IOException, ExcelDocumentNotSupported
	{
		if (filePath.endsWith(".xlsx") || !filePath.contains("."))
		{
			return new com.genexus.office.poi.xssf.ExcelSpreadsheet(handler, filePath, template);
		}
		throw new ExcelDocumentNotSupported();
	}

}
