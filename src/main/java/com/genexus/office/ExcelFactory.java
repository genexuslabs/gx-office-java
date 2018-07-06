package com.genexus.office;

import java.io.IOException;

import com.genexus.gxoffice.IGxError;
import com.genexus.office.exception.ExcelDocumentNotSupported;
import com.genexus.office.exception.ExcelTemplateNotFoundException;

public class ExcelFactory
{

	public static IExcelSpreadsheet create(IGxError handler, String filePath, String template)
			throws ExcelTemplateNotFoundException, IOException, ExcelDocumentNotSupported
	{
		if (filePath.endsWith(".xlsx") || !filePath.contains("."))
		{
			return new com.genexus.office.poi.xssf.ExcelSpreadsheet(handler, filePath, template);
		}
		throw new ExcelDocumentNotSupported();
	}

}
