package com.genexus.office.exception;

import com.genexus.office.ErrorCodes;

public class ExcelDocumentNotSupported extends ExcelException
{
	public ExcelDocumentNotSupported()
	{
		super(com.genexus.office.ErrorCodes.EXTENSION_NOT_SUPPORTED, "File extension not supported");
	}
}
