package com.genexus.office;

public interface IExcelWorksheet
{
	public String getName();

	public Boolean isHidden();

	public Boolean rename(String newName);
	
	public Boolean copy(String newName);

	public void setProtected(String password);
}
