package com.genexus.office.exception;

public class ExcelException extends Throwable
{

	private int _errorCode;
	private String _errDsc;

	public ExcelException(int errCode, String errDsc)
	{
		super();
		_errorCode = errCode;
		_errDsc = errDsc;
	}

	public int get_errorCode()
	{
		return _errorCode;
	}

	public String get_errDsc()
	{
		return _errDsc;
	}

}