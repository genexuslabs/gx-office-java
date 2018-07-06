package com.genexus.office.exception;

import com.genexus.office.ErrorCodes;

public class ExcelReadonlyException extends ExcelException
{
	public ExcelReadonlyException()
	{
		super(13, "Can not modify a readonly document");
	}
}
