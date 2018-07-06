package com.genexus.office.exception;

import com.genexus.office.ErrorCodes;

public class ExcelTemplateNotFoundException extends ExcelException
{
	public ExcelTemplateNotFoundException()
	{
		super(ErrorCodes.TEMPLATE_NOT_FOUND, "Template not found");
	}
}
