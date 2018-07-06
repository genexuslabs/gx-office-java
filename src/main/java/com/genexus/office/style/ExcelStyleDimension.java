package com.genexus.office.style;

public abstract class ExcelStyleDimension {
	
	private boolean isDirty = false;
	
	public boolean isDirty() {
		return isDirty;
	}
	
	public void setChanged() {
		isDirty = true;
	}
}
